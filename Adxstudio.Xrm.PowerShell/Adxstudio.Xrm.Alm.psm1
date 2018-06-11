function New-AlmServerSetting
{
	<#
	.SYNOPSIS
	Defines the settings for a Dynamics CRM on-premises server. Pass the result to the 'New-AlmContext' function in order to construct a full ALM context.

	.EXAMPLE
	$crm2013CrmUrl = "https://internal.crm2013.contoso.com"
	$crm2013SqlServerName = "crm2013"
	$crm2013SrsUrl = "http://crm2013/reportserver"
	$credential = Get-Credential -Username "CONTOSO\CrmDeploymentAdministrator" -Message "Sign-in as a CRM deployment administrator."
	$isActiveDirectory = $true
	$useIntegratedTfsAuthentication = $true
	$useIntegratedSqlAuthentication = $true
	$sqlBackupFile = "C:\temp\New_MSCRM.bak"
	$sqlServer2012DataPath = "C:\Program Files\Microsoft SQL Server\MSSQL11.MSSQLSERVER\MSSQL\DATA\"
	$crm2013 = New-AlmServerSetting $crm2013CrmUrl $crm2013SqlServerName $crm2013SrsUrl $credential -IsActiveDirectory:$isActiveDirectory -UseIntegratedTfsAuthentication:$useIntegratedTfsAuthentication -TfsCredential $tfsCredential -UseIntegratedSqlAuthentication:$useIntegratedSqlAuthentication -SqlCredential $sqlCredential -SqlBackupFile $sqlBackupFile -SqlDataPath $sqlServer2012DataPath
	
	$demoSolutions = New-AlmSolutionSetting ...
	$demoData = New-AlmDataSetting ...
	$dataEncryptionKey = "??????????????????????????????"
	$crmOrg = New-AlmOrganizationSetting "adventureworks1" "Adventure Works 1" $demoSolutions $demoData -DataEncryptionKey $dataEncryptionKey
	$almContext = New-AlmContext "CRM/aw&1" $crm2013 $crmOrg -Force:$force
	#>

	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true)]
		$deploymentServiceUrl,

		[parameter(Position=1, Mandatory=$true)]
		$sqlServerName,

		[parameter(Position=2, Mandatory=$true)]
		$srsUrl,

		[parameter(Position=3)]
		$credential,

		$timeout = "01:00:00",

		$defaultUsername = "CONTOSO\administrator",

		[Switch]
		$useIntegratedSqlAuthentication,

		$sqlCredential,

		$sqlDefaultUsername = "administrator",

		$sqlBackupFile,

		$sqlDataPath = "C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\",

		[Switch]
		$useIntegratedTfsAuthentication,

		$tfsCredential,
 
		$getCredentialMessage = "Sign-in as a CRM deployment administrator.",

		$getSqlCredentialMessage = "Sign-in as a SQL Server authentication user.",

		$getTfsCredentialMessage = "Sign-in as a TFS user.",

		[switch]
		$isActiveDirectory
	)

	function Get-Cred
	{
		param ([parameter(Position=0)]$username, [parameter(Position=1)]$message)

		if ($PSVersionTable.PSVersion.Major -ge 3)
		{
			return Get-Credential -UserName $username -Message $message
		}
		else
		{
			Write-Host $message
			return Get-Credential $username
		}
	}

	if (-not $credential)
	{
		$credential = Get-Cred $defaultUsername $getCredentialMessage
	}

	if ((-not $sqlCredential) -and (-not $useIntegratedSqlAuthentication))
	{
		# collect SQL Server authentication mode based credentials

		$sqlCredential = Get-Cred $sqlDefaultUsername $getSqlCredentialMessage
	}

	if ((-not $tfsCredential) -and (-not $useIntegratedTfsAuthentication))
	{
		$tfsCredential = Get-Cred $defaultUsername $getTfsCredentialMessage
	}

	$deployConnection = Get-CrmConnection -Url $deploymentServiceUrl -Credential $credential -Timeout $timeout
	$discoConnection = Get-CrmConnection -Url $deploymentServiceUrl -Credential $credential -Timeout $timeout

	return @{
		DeploymentServiceConnection = $deployConnection;
		DiscoveryServiceConnection = $discoConnection;
		DeploymentServiceUrl = $deploymentServiceUrl;
		SqlServerName = $sqlServerName;
		SrsUrl = $srsUrl;
		Credential = $credential;
		Timeout = $timeout;
		DefaultUsername = $defaultUsername;
		UseSqlServerAuthentication = $useSqlServerAuthentication;
		SqlCredential = $sqlCredential;
		SqlDefaultUsername = $sqlDefaultUsername;
		SqlDataPath = $sqlDataPath;
		SqlBackupFile = $sqlBackupFile;
		GetCredentialMessage = $getCredentialMessage;
		TfsCredential = $tfsCredential;
		IsActiveDirectory = $isActiveDirectory;
	} | ConvertTo-PSObject
}

function New-AlmCrmOnlineServerSetting
{
	<#
	.SYNOPSIS
	Defines the settings for a Dynamics CRM Online server. Pass the result to the 'New-AlmContext' function in order to construct a full ALM context.

	.EXAMPLE
	$crmOnline = New-AlmCrmOnlineServerSetting -defaultUsername $defaultUsernameOnline                    # prompts for user credentials
	$crmOnlineOrg = Choose-CrmOrganization $crmOnline                                                     # prompts for the organization
	$crmOnlineOrgSetting = New-AlmOrganizationSetting $crmOnlineOrg.UniqueName $crmOnlineOrg.FriendlyName # set the organization unique name and display name
	$almContext = New-AlmContext "CRM Online" $crmOnline $crmOnlineOrgSetting -ResolveUrl                 # create the ALM context and resolve the org service URL
	#>

	[CmdletBinding()]
	param (
		$deploymentServiceUrl,

		[parameter(Position=0)]
		$credential,

		[parameter(Position=1)]
		$timeout = "01:00:00",

		$defaultUsername = "administrator@contoso.onmicrosoft.com",

		$getCredentialMessage = "Sign-in as a CRM deployment administrator.",

		[switch]
		$emea,

		[switch]
		$apac
	)

	if (-not $deploymentServiceUrl) {
		if ($emea) {
			$deploymentServiceUrl = "https://disco.crm4.dynamics.com/"
		} elseif ($apac) {
			$deploymentServiceUrl = "https://disco.crm5.dynamics.com/"
		} else {
			$deploymentServiceUrl = "https://disco.crm.dynamics.com/"
		}
	}

	$emptyCred = New-Object System.Management.Automation.PSCredential("username", (" " | ConvertTo-SecureString -AsPlainText -Force))
	return (New-AlmServerSetting $deploymentServiceUrl "" "" $credential -timeout $timeout -sqlCredential $emptyCred -tfsCredential $emptyCred -defaultUsername $defaultUsername -getCredentialMessage $getCredentialMessage)
}

function New-AlmSolutionSetting
{
	<#
	.SYNOPSIS
	Defines the settings for a single solution resource. Pass the result to the 'New-AlmOrganizationSetting' function in order to construct a full ALM context.

	.EXAMPLE
	# specify additional versioned solution export files
	# http://msdn.microsoft.com/en-us/library/dn689055.aspx
	$targetVersion = (
		@{ CustomizationFile = "$scriptPath\Customizations\MySolution01_target_CRM_5.0.zip"; TargetVersion = "5.0.0.0" },
		@{ CustomizationFile = "$scriptPath\Customizations\MySolution01_target_CRM_6.0.zip"; TargetVersion = "6.0.0.0" },
		@{ CustomizationFile = "$scriptPath\Customizations\MySolution01_managed_target_CRM_5.0.zip"; TargetVersion = "5.0.0.0"; Managed = $true },
		@{ CustomizationFile = "$scriptPath\Customizations\MySolution01_managed_target_CRM_6.0.zip"; TargetVersion = "6.0.0.0"; Managed = $true }
	) | ConvertTo-PSObject

	$demoSolutions = (
		(New-AlmSolutionSetting "MySolution01" "$scriptPath\Customizations\MySolution01.zip" -Resolve -TargetVersion $targetVersion),
		(New-AlmSolutionSetting "MySolution02" "$scriptPath\Customizations\MySolution02.zip" -Resolve)
	)

	$crm2013 = New-AlmServerSetting ...
	$demoData = New-AlmDataSetting ...
	$dataEncryptionKey = "??????????????????????????????"
	$crmOrg = New-AlmOrganizationSetting "adventureworks1" "Adventure Works 1" $demoSolutions $demoData -DataEncryptionKey $dataEncryptionKey
	$almContext = New-AlmContext "CRM/aw&1" $crm2013 $crmOrg -Force:$force
	#>

	[CmdletBinding()]
	param(
		[parameter(Position=0, Mandatory=$true)]
		$uniqueName,

		[parameter(Position=1, Mandatory=$true)]
		$customizationFile,

		[parameter(Position=2)]
		$managedFile,

		$targetVersion,

		[switch]
		$managed,

		[switch]
		$resolve
	)

	# validate the source files

	if ($resolve -and $customizationFile)
	{
		$resolvedCustomizationFile = $customizationFile | Resolve-Path | % { $_.Path }

		if (-not $resolvedCustomizationFile)
		{
			throw "Failed to load solution files."
		}
	}
	else
	{
		$resolvedCustomizationFile = $customizationFile
	}

	return @{
		UniqueName = $uniqueName;
		CustomizationFile = $resolvedCustomizationFile;
		ManagedFile = $managedFile;
		Managed = $managed;
		TargetVersion = $targetVersion;
	} | ConvertTo-PSObject
}

function New-AlmDataSetting
{
	<#
	.SYNOPSIS
	Defines the settings for a single data resource. Pass the result to the 'New-AlmOrganizationSetting' function in order to construct a full ALM context.

	.EXAMPLE
	$demoData = New-AlmDataSetting "Default" "$scriptPath\Data\Dev" -Generalized -Uncompressed -Resolve

	$crm2013 = New-AlmServerSetting ...
	$demoSolutions = New-AlmSolutionSetting ...
	$dataEncryptionKey = "??????????????????????????????"
	$crmOrg = New-AlmOrganizationSetting "adventureworks1" "Adventure Works 1" $demoSolutions $demoData -DataEncryptionKey $dataEncryptionKey
	$almContext = New-AlmContext "CRM/aw&1" $crm2013 $crmOrg -Force:$force
	#>

	[CmdletBinding()]
	param(
		[parameter(Position=0, Mandatory=$true)]
		$name,

		[parameter(Position=1, Mandatory=$true)]
		$dataFile,

		$dataPackageFile,

		$contentFetchXml,

		[string[]]
		$exclude,

		[string[]]
		$include,

		[Hashtable]
		$attributesToExclude,

		[int]
		$pageCount,

		[switch]
		$uncompressed,

		[switch]
		$excludeMetadata,

		[switch]
		$excludeSystem,

		[switch]
		$generalized,

		[switch]
		$excludeDefaultValues,

		[switch]
		$insertNullValueAttributes,

		[switch]
		$insertDefaultValueAttributes,

		[switch]
		$resolve
	)

	# validate the source files

	if ($resolve -and $dataFile)
	{
		$resolvedDataFile = $dataFile | Resolve-Path | % { $_.Path }

		if (-not $resolvedDataFile)
		{
			throw "Failed to load data file."
		}
	}
	else
	{
		$resolvedDataFile = $dataFile
	}

	return @{
		Name = $name;
		DataFile = $resolvedDataFile;
		DataPackageFile = $dataPackageFile;
		ExportContentParams = @{
			ContentFetchXml = $contentFetchXml;
			Exclude = $exclude;
			Include = $include;
			AttributesToExclude = $attributesToExclude;
			PageCount = $pageCount;
			Uncompressed = $uncompressed;
			Generalized = $generalized;
			ExcludeDefaultValues = $excludeDefaultValues;
			InsertNullValueAttributes = $insertNullValueAttributes;
			InsertDefaultValueAttributes = $insertDefaultValueAttributes;
			ExcludeMetadata = $excludeMetadata;
			ExcludeSystem = $excludeSystem;
		};
	} | ConvertTo-PSObject
}

function New-AlmOrganizationSetting
{
	<#
	.SYNOPSIS
	Defines the settings for a CRM organization. Pass the result to the 'New-AlmContext' function in order to construct a full ALM context.

	.EXAMPLE
	$crm2013 = New-AlmServerSetting ...
	$demoSolutions = New-AlmSolutionSetting ...
	$demoData = New-AlmDataSetting ...
	$dataEncryptionKey = "??????????????????????????????"
	$crmOrg = New-AlmOrganizationSetting "adventureworks1" "Adventure Works 1" $demoSolutions $demoData -DataEncryptionKey $dataEncryptionKey
	$almContext = New-AlmContext "CRM/aw&1" $crm2013 $crmOrg -Force:$force
	#>

	[CmdletBinding()]
	param(
		[parameter(Position=0, Mandatory=$true)]
		$organizationName,

		[parameter(Position=1)]
		$displayName,

		[parameter(Position=2)]
		$solution,

		[parameter(Position=3)]
		$data,

		$currency,

		[int]
		$baseLanguageCode,

		[int[]]
		$languageCode,

		$dataEncryptionKey,

		$preImportStatePath,

		[System.Collections.Hashtable[]]
		$newUsers
	)

	return @{
		OrganizationName = $organizationName;
		DisplayName = $displayName;
		Solution = $solution;
		Data = $data;
		Currency = $currency;
		BaseLanguageCode = $baseLanguageCode;
		LanguageCode = $languageCode;
		DataEncryptionKey = $dataEncryptionKey;
		PreImportStatePath = $preImportStatePath;
		NewUsers = $newUsers;
	} | ConvertTo-PSObject
}

function New-AlmContext
{
	<#
	.SYNOPSIS
	Defines the complete set of settings for a single CRM organization.

	.EXAMPLE
	$crm2013 = New-AlmServerSetting ...
	$demoSolutions = New-AlmSolutionSetting ...
	$demoData = New-AlmDataSetting ...
	$dataEncryptionKey = "??????????????????????????????"
	$crmOrg = New-AlmOrganizationSetting "adventureworks1" "Adventure Works 1" $demoSolutions $demoData -DataEncryptionKey $dataEncryptionKey
	$almContext = New-AlmContext "CRM/aw&1" $crm2013 $crmOrg -Force:$force
	#>

	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true)]
		$name,

		[parameter(Position=1, Mandatory=$true)]
		$server,

		[parameter(Position=2)]
		$organization,

		[parameter(Position=3)]
		$organizationServiceUrl,

		[Switch]
		$force,

		[switch]
		$resolveUrl
	)

	# create the connections

	if ($organizationServiceUrl) {
		# directly use the explicit organization service URL

		$orgUrl = $organizationServiceUrl
	}
	elseif ($organization.OrganizationName) {
		if ($resolveUrl -or (-not $server.IsActiveDirectory)) {
			# resolve the organization service URL with the discovery service

			$orgs = Get-CrmOrganization -Connection $server.DiscoveryServiceConnection
			$org = $orgs | ? { $_.UniqueName -eq $organization.OrganizationName }

			if (-not $org) {
				throw ("The organization '{0}' was not found.", $organization.OrganizationName)
			}

			$orgUrl = $org.Endpoints["Web"]
		} else {
			# fallback to the internal organization service URL format

			$orgUrl = "{0}/{1}" -f $server.DeploymentServiceUrl, $organization.OrganizationName
		}
	}
	else
	{
		$orgUrl = $null
	}

	$deployConnection = $server.DeploymentServiceConnection
	$discoConnection = $server.DeploymentServiceConnection

	if ($orgUrl)
	{
		$orgConnection = Get-CrmConnection -Url $orgUrl -Credential $server.Credential -Timeout $server.Timeout
	}
	else
	{
		$orgConnection = $null
	}

	if ($organization -and -not $organization.DisplayName)
	{
		$organization.DisplayName = $organization.OrganizationName
	}

	$customizationFile = $organization.Solution | % { $_.CustomizationFile }
	$dataFile = $organization.Data | % { $_.DataFile }

	$cmdletMask = "*-Crm*"
	$cmdletMaskSql = "*-Sql*"
	$cmdletMaskTfs = "*-Tfs*"
	$cmdletMaskAlm = "*-Alm*"

	$defaultParameterValues = @{
		"${cmdletMask}:deploymentServiceUrl" = $server.DeploymentServiceUrl;
		"${cmdletMask}:discoveryServiceUrl" = $server.DeploymentServiceUrl;
		"${cmdletMask}:organizationServiceUrl" = $orgUrl;

		"${cmdletMask}:deploymentServiceConnection" = $deployConnection;
		"${cmdletMask}:discoveryServiceConnection" = $discoConnection;
		"${cmdletMask}:organizationServiceConnection" = $orgConnection;

		"${cmdletMask}:credential" = $server.Credential;
		"${cmdletMask}:sqlCredential" = $server.SqlCredential;
		"${cmdletMaskSql}:sqlCredential" = $server.SqlCredential;
		"${cmdletMask}:tfsCredential" = $server.TfsCredential;
		"${cmdletMaskTfs}:tfsCredential" = $server.TfsCredential;

		"${cmdletMask}:sqlServerName" = $server.SqlServerName;
		"${cmdletMaskSql}:sqlServerName" = $server.SqlServerName;
		"${cmdletMask}:srsUrl" = $server.SrsUrl;

		"${cmdletMask}:force" = $force;

		"${cmdletMask}:organizationName" = $organization.OrganizationName;
		"${cmdletMask}:sqlDataPath" = $server.SqlDataPath;
		"${cmdletMask}:sqlBackupFile" = $server.SqlBackupFile;
		"${cmdletMask}:displayName" = $organization.DisplayName;
		"${cmdletMask}:currency" = $organization.Currency;
		"${cmdletMask}:baseLanguageCode" = $organization.BaseLanguageCode;
		"${cmdletMask}:languageCode" = $organization.LanguageCode;
		"${cmdletMask}:customizationFile" = $customizationFile;
		"${cmdletMask}:dataFile" = $dataFile;
		"${cmdletMask}:newUsers" = $organization.NewUsers;
	}

	return @{
		Name = $name;
		Force = $force;

		Server = $server;
		Organization = $organization;

		DeploymentServiceConnection = $deployConnection;
		DiscoveryServiceConnection = $discoConnection;
		OrganizationServiceConnection = $orgConnection;

		DefaultParameterValues = $defaultParameterValues;
	} | ConvertTo-PSObject
}

function Write-AlmDefaultParameterValues
{
	<#
	.SYNOPSIS
	Displays the currently defined default parameter values.

	.EXAMPLE
	$almContext = New-AlmContext ...
	$PSDefaultParameterValues = $almContext.DefaultParameterValues
	Write-AlmDefaultParameterValues $PSDefaultParameterValues
	#>

	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[hashtable]
		$values
	)

	Write-Host "------------------------"
	Write-Host "Default Parameter Values"
	Write-Host "------------------------"

	$values.GetEnumerator() | Sort-Object Name
}

function Choose-CrmConnection
{
	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		$settings,

		$caption = "Select Connection",
		$message = "Select a connection."
	)

	process {
		$choices = $settings | % { New-Object System.Management.Automation.Host.ChoiceDescription($_.Name, ("{0}/{1}" -f $_.Server.DeploymentServiceUrl, $_.Organization.OrganizationName)) }
		$options = [System.Management.Automation.Host.ChoiceDescription[]]($choices)
		$result = $host.ui.PromptForChoice($caption, $message, $options, 0)
		$setting = $settings[$result]

		return $setting
	}
}

function Choose-CrmOrganization
{
	<#
	.SYNOPSIS
	Choose from a set of existing organizations on a single CRM server. Suitable for reading and updating organizations.

	.EXAMPLE
	$crm2013 = New-AlmServerSetting ...
	$org = Choose-CrmOrganization $crm2013
	$orgSetting = New-AlmOrganizationSetting $org.UniqueName $org.FriendlyName
	$almContext = New-AlmContext ("CRM2013/{0}" -f $org.UniqueName) $crm2013 $orgSetting -Force:$force
	#>

	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		$server,

		$caption,
		$message = "Select an organization.",

		[switch]
		$allowAll
	)

	process {
		$orgs = Get-CrmOrganization -Connection $server.DiscoveryServiceConnection | Sort-Object UniqueName
		return Choose-AlmCollection $orgs "UniqueName" "FriendlyName" -Caption $caption -Message $message -AllowAll:$allowAll
	}
}

function Choose-AlmCollection
{
	<#
	.SYNOPSIS
	Helper function for displaying a multiple choice prompt.

	.EXAMPLE
	$crmServers = (
		@{ Label = "CRM 2013" },
		@{ Label = "CRM Online" }
	) | ConvertTo-PSObject

	$crmServer = Choose-AlmCollection $crmServers "Label" "Label" -message "Select a server."
	#>

	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		$collection,

		[parameter(Position=1, Mandatory=$true)]
		$labelProperty,

		[parameter(Position=2, Mandatory=$true)]
		$descriptionProperty,

		$caption,
		$message = "Select",

		[switch]
		$allowAll,

		[switch]
		$excludeIndex
	)

	process {
		$indexed = ConvertTo-IndexedPair $collection
		$choices = $indexed | % {
			if (-not $excludeIndex) {
				$index = $_.Key + 1
				if ($index -lt 10) {
					$hotKey = ("&{0}) " -f $index)
				} elseif ($index -lt 36) {
					# convert to alphabet
					$hotKey = ("&{0}) " -f [char]($index + 55))
				} else {
					$hotKey = $null
				}
			}
			New-Object System.Management.Automation.Host.ChoiceDescription(("{0}{1}" -f $hotKey, (Select-Object -InputObject $_.Value -ExpandProperty $labelProperty)), (Select-Object -InputObject $_.Value -ExpandProperty $descriptionProperty))
		}

		if ($allowAll) {
			$choices += (New-Object System.Management.Automation.Host.ChoiceDescription("&*) All"))
		}

		$options = [System.Management.Automation.Host.ChoiceDescription[]]($choices)
		$serverCaption = ($caption, $server.DeploymentServiceUrl -ne $null)[0]
		$result = $host.ui.PromptForChoice($serverCaption, $message, $options, 0)
		
		if ($result -eq $collection.Length) {
			$item = $collection
		} else {
			$item = $collection[$result]
		}

		return $item
	}
}

function Provision-AlmDeployment
{
	<#
	.SYNOPSIS
	Provisions an organization defined by the ALM context.

	.EXAMPLE
	$almContext = New-AlmContext ...
	Provision-AlmDeployment $almContext -RemoveExistingOrganization
	#>

	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		$almContext,

		[switch]
		$ignoreSqlBackupFile,

		[switch]
		$removeExistingOrganization,

		[switch]
		$importSolutionAsJob,

		$importSolutionTimeout = "00:30:00"
	)

	process
	{
		if (-not $almContext.Organization)
		{
			throw "Provide a connection that specifies an organization."
		}

		$isAdmin = Test-CrmIsDeploymentAdmin -Connection $almContext.DeploymentServiceConnection

		if (-not $isAdmin)
		{
			throw "Provision terminated! The user for the provided credentials is not a Deployment Administrator."
		}

		if ($removeExistingOrganization)
		{
			# remove the existing organization

			$entityInstanceId = Get-CrmEntityInstanceId -Name $almContext.Organization.OrganizationName

			Remove-CrmDeployment -Verbose -Connection $almContext.DeploymentServiceConnection -SqlCredential $almContext.Server.SqlCredential -EntityInstanceId $entityInstanceId -Force:$almContext.Force
		}

		if ((-not $ignoreSqlBackupFile) -and $almContext.Server.SqlBackupFile)
		{
			# restore an existing SQL database backup for importing the new organization

			$sqlDatabaseName = Restore-CrmSqlDatabase -SqlServerName $almContext.Server.SqlServerName -SqlCredential $almContext.Server.SqlCredential -OrganizationName $almContext.Organization.OrganizationName -SqlBackupFile $almContext.Server.SqlBackupFile -SqlDataPath $almContext.Server.SqlDataPath

			Write-Verbose ("Database restored: {0}" -f $sqlDatabaseName) -verbose
		}

		if (-not $almContext.Server.IsActiveDirectory) {
			Write-Warning "*** For IFD based connections, it may be necessary to update the relying party trust after the organization"
			Write-Warning "is created for the first time. ***"
			Write-Warning " - See: http://technet.microsoft.com/en-us/library/ee892339.aspx"
			Write-Warning " - Example: Update-ADFSRelyingPartyTrust -TargetName auth.crm2013.contoso.com"
		}

		$customizationFile = $almContext.Organization.Solution | % { $_.CustomizationFile }
		$dataFile = $almContext.Organization.Data | % { $_.DataFile }

		$orgUrl = New-CrmDeployment -Verbose -SqlDatabaseName $sqlDatabaseName -Connection $almContext.DeploymentServiceConnection -OrganizationName $almContext.Organization.OrganizationName -DisplayName $almContext.Organization.DisplayName -SqlServerName $almContext.Server.SqlServerName -SrsUrl $almContext.Server.SrsUrl -Currency $almContext.Organization.Currency -NewUsers $almContext.Organization.NewUsers -DataEncryptionKey $almContext.Organization.DataEncryptionKey -PreImportStatePath $almContext.Organization.PreImportStatePath -CustomizationFile $customizationFile -DataFile $dataFile -ImportSolutionAsJob:$importSolutionAsJob -ImportSolutionTimeout $importSolutionTimeout

		Write-Output $orgUrl
	}
}

function Remove-AlmDeployment
{
	<#
	.SYNOPSIS
	Removes an organization defined by the ALM context.

	.EXAMPLE
	$almContext = New-AlmContext ...
	Remove-AlmDeployment $almContext
	#>

	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		$almContext
	)

	process
	{
		if (-not $almContext.Organization)
		{
			throw "Provide a connection that specifies an organization."
		}

		$isAdmin = Test-CrmIsDeploymentAdmin -Connection $almContext.DeploymentServiceConnection

		if (-not $isAdmin)
		{
			throw "Provision terminated! The user for the provided credentials is not a Deployment Administrator."
		}

		$entityInstanceId = Get-CrmEntityInstanceId $almContext.Organization.OrganizationName

		Remove-CrmDeployment -Verbose -Connection $almContext.DeploymentServiceConnection -SqlCredential $almContext.Server.SqlCredential -EntityInstanceId $entityInstanceId -Force:$almContext.Force
	}
}

function Update-AlmDeployment
{
	<#
	.SYNOPSIS
	Updates an organization defined by the ALM context.

	.EXAMPLE
	$almContext = New-AlmContext ...
	Update-AlmDeployment $almContext
	#>

	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		$almContext,

		[switch]
		$excludeCustomization,

		[switch]
		$excludeData,

		[switch]
		$alwaysImport,

		[switch]
		$importSolutionAsJob,

		$importSolutionTimeout = "00:30:00",

		$updatesEnabled = $true
	)

	process
	{
		if (-not $almContext.Organization)
		{
			throw "Provide a connection that specifies an organization."
		}

		if (-not $excludeCustomization)
		{
			# get the version values for the currently imported solutions

			$existingSolutions = Get-CrmSolutionVersion -Connection $almContext.OrganizationServiceConnection | ConvertTo-PSObject

			Write-Host "Current solutions:"
			$existingSolutions | ft

			# iterate over the solutions passed in from the file system

			$customizationFile = $almContext.Organization.Solution | % { $_.CustomizationFile }
			$customizations = $customizationFile | Get-CrmSolutionVersion | ConvertTo-PSObject

			$customizations | % {

				$customization = $_

				Write-Host "Comparing customization:"
				$customization | ft

				# find the matching imported solution

				$existingSolution = ($existingSolutions | Where-Object { $_.uniquename -eq $customization.uniquename } | Select-Object -First 1)

				# determine if the version of the customization file is newer (greater) than the existing solution

				$newVersion = ($existingSolution | Where-Object { $_.version -lt $customization.version })
	
				# either the customization is a newer version or was never imported

				$filename = $customization.filename

				if ($alwaysImport -or (-not $existingSolution) -or $newVersion)
				{
					Write-Host "Importing solution: $filename"

					if ($importSolutionAsJob)
					{
						$importJob = Import-CrmSolutionAsJob $almContext.OrganizationServiceConnection $filename -ImportSolutionTimeout $importSolutionTimeout
					}
					else
					{
						$importJob = Import-CrmSolutionSynchronous $almContext.OrganizationServiceConnection $filename
					}

					Write-Host "Imported solution:"
					$importJob | ft | Out-Host
				}
				else
				{
					Write-Warning "Skipped import solution: $filename"

					if (-not $newVersion)
					{
						Write-Warning "Version of requested solution import is not newer than the existing solution version."
					}
				}
			}
		}

		if (-not $excludeData)
		{
			$almContext.Organization.Data | % {
				$data = $_
				Import-CrmContent -Connection $almContext.OrganizationServiceConnection -InputPath $data.DataFile -PreImportStatePath $almContext.Organization.PreImportStatePath -UpdatesEnabled:$updatesEnabled -Force
			}
		}
	}
}

function Export-AlmDeployment
{
	<#
	.SYNOPSIS
	Exports an organization defined by the ALM context.

	.EXAMPLE
	$almContext = New-AlmContext ...
	Export-AlmDeployment $almContext
	#>

	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		$almContext,

		$customizationCheckinComment,

		$dataCheckinComment,

		[scriptblock]
		$transform,

		[switch]
		$excludeCustomization,

		[switch]
		$excludeData,

		[switch]
		$checkout,

		[switch]
		$checkin
	)

	if (-not $excludeCustomization)
	{
		# export solutions

		$almContext.Organization.Solution | % {
			if (-not $_.Managed) {
				# only export unmanaged solutions

				$solutionName = $_.UniqueName

				if ($_.TargetVersion -is [string]) {
					$baseTargetVersion = $_.TargetVersion -as [string]
				} elseif ($_.TargetVersion -is [array]) {
					$additionalTargetVersion = $_.TargetVersion -as [array]
				}

				if ($_.CustomizationFile) {
					Export-AlmCustomizationFile $almContext -customizationFile $_.CustomizationFile -UniqueName $solutionName -Managed:$false -TargetVersion $baseTargetVersion -checkin:$checkin -checkout:$checkout -Transform $transform
				}

				if ($_.ManagedFile) {
					# also export a managed version of the unmanaged solution

					Export-AlmCustomizationFile $almContext -customizationFile $_.ManagedFile -UniqueName $solutionName -Managed:$true -TargetVersion $baseTargetVersion -checkin:$checkin -checkout:$checkout -Transform $transform
				}

				if ($additionalTargetVersion) {
					# export additional versioned unmanaged or managed solutions

					$additionalTargetVersion | % {
						Export-AlmCustomizationFile $almContext -customizationFile $_.CustomizationFile -UniqueName $solutionName -Managed:$_.Managed -TargetVersion $_.TargetVersion -checkin:$checkin -checkout:$checkout -Transform $transform
					}
				}
			}
		}
	}

	if (-not $excludeData)
	{
		# export data

		$almContext.Organization.Data | % {
			$data = $_

			if ($checkout)
			{
				$dataExists = Test-TfsPath -Credential $almContext.Server.TfsCredential $data.DataFile

				if ($dataExists)
				{
					Invoke-TfsCheckout -Credential $almContext.Server.TfsCredential $data.DataFile -Recursive:$data.ExportContentParams.Uncompressed
				}
			}

			Write-Host ("Exporting {0} data..." -f $data.Name)

			$exportCrmContentParams = $data.ExportContentParams.Clone()
			$exportCrmContentParams["Connection"] = $almContext.OrganizationServiceConnection
			$exportCrmContentParams["OutputPath"] = $data.DataFile

			Export-CrmContent @exportCrmContentParams

			if ($data.DataPackageFile) {
				Merge-CrmContent $data.DataFile -ContentOutputPath $data.DataPackageFile | % { "Written to: {0}" -f $_ }
			}

			if ($checkin)
			{
				if (-not $dataExists)
				{
					Invoke-TfsAdd -Credential $almContext.Server.TfsCredential $data.DataFile -Recursive:$data.ExportContentParams.Uncompressed
				}

				Invoke-TfsCheckin -Credential $almContext.Server.TfsCredential $data.DataFile -Recursive:$data.ExportContentParams.Uncompressed -Comment $dataCheckinComment
			}
			elseif ($checkout)
			{
				Write-Host ("Review pending change to: {0}" -f $data.DataFile)
			}
		}
	}
}

function Export-AlmCustomizationFile {
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		$almContext,

		$customizationFile,

		$uniqueName,

		[string]
		$targetVersion,

		[scriptblock]
		$transform,

		[switch]
		$managed,

		[switch]
		$checkout,

		[switch]
		$checkin
	)

	if ($checkout)
	{
		$customizationExists = Test-TfsPath -Credential $almContext.Server.TfsCredential $customizationFile

		if ($customizationExists)
		{
			Invoke-TfsCheckout -Credential $almContext.Server.TfsCredential $customizationFile
		}
	}

	Write-Host ("Exporting {0} solution..." -f $uniqueName)

	if ($transform) {
		Export-CrmSolution -Connection $almContext.OrganizationServiceConnection -UniqueName $uniqueName -TargetVersion $targetVersion -Managed:$managed | % {
			Write-Host ("Transforming {0} solution..." -f $uniqueName)

			$solutionEntity = $_.Key
			$solutionFile = $_.Value
			$transformedFile = & $transform -Entity $solutionEntity -File $solutionFile
			[System.IO.File]::WriteAllBytes($customizationFile, $transformedFile)

			Write-Output $_
		}
	} else {
		Export-CrmSolution -Connection $almContext.OrganizationServiceConnection -UniqueName $uniqueName -TargetVersion $targetVersion -OutputPath $customizationFile -Managed:$managed
	}

	if ($checkin)
	{
		if (-not $customizationExists)
		{
			Invoke-TfsAdd -Credential $almContext.Server.TfsCredential $customizationFile
		}

		Invoke-TfsCheckin -Credential $almContext.Server.TfsCredential $customizationFile -Comment $customizationCheckinComment
	}
	elseif ($checkout)
	{
		Write-Host ("Review pending change to: {0}" -f $customizationFile)
	}
}

function ConvertTo-AlmSolutionXml
{
	param (
		[parameter(position=0,Mandatory=$true,ValueFromPipeline=$true)]
		[System.IO.Packaging.PackagePart]
		$part
	)

	Invoke-Using ([System.IO.Stream] $stream = $part.GetStream([System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)) {
		Write-Output ([System.Xml.Linq.XElement]::Load($stream))
	}
}

function Replace-AlmSolutionPart
{
	param (
		[parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[System.Xml.Linq.XElement]
		$xml,

		[parameter(position=0)]
		[System.IO.Packaging.Package]
		$package,

		[parameter(position=1)]
		[System.IO.Packaging.PackagePart]
		$part
	)

	$uri = $part.Uri
	$contentType = $part.ContentType
	$compressionOption = $part.CompressionOption
	$package.DeletePart($part.Uri)
	$replacement = $package.CreatePart($uri, $contentType, $compressionOption)

	Invoke-Using ([System.IO.Stream] $stream = $replacement.GetStream()) {
		$xml.Save($stream)
	}
}

function ConvertTo-AlmSolutionTransformed
{
	param (
		[parameter(position=0,Mandatory=$true,ValueFromPipeline=$true)]
		[byte[]]
		$file,

		[parameter(position=1)]
		[scriptblock]
		$transformCustomizations,

		[parameter(position=2)]
		[scriptblock]
		$transformSolution
	)

	Invoke-Using ($stream = New-Object System.IO.MemoryStream) {
		$stream.Write($file, 0, $file.Length)
		Invoke-Using ($package = [System.IO.Packaging.Package]::Open($stream, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite)) {
			$parts = $package.GetParts();

			if ($transformCustomizations) {
				$customizationsPart = $parts | ? { $_.Uri.OriginalString -eq "/customizations.xml" }
				$customizationsPart | ConvertTo-AlmSolutionXml | % { & $transformCustomizations -Xml $_ } | Replace-AlmSolutionPart -Package $package -Part $customizationsPart
			}

			if ($transformSolution) {
				$solutionPart = $parts | ? { $_.Uri.OriginalString -eq "/solution.xml" }
				$solutionPart | ConvertTo-AlmSolutionXml | % { & $transformSolution -Xml $_ } | Replace-AlmSolutionPart -Package $package -Part $solutionPart
			}
		}
		Write-Output $stream.ToArray()
	}
}

Export-ModuleMember -function New-AlmServerSetting
Export-ModuleMember -function New-AlmCrmOnlineServerSetting
Export-ModuleMember -function New-AlmSolutionSetting
Export-ModuleMember -function New-AlmDataSetting
Export-ModuleMember -function New-AlmOrganizationSetting
Export-ModuleMember -function New-AlmContext
Export-ModuleMember -function Write-AlmDefaultParameterValues
Export-ModuleMember -function Choose-CrmConnection
Export-ModuleMember -function Choose-CrmOrganization
Export-ModuleMember -function Choose-AlmCollection
Export-ModuleMember -function Provision-AlmDeployment
Export-ModuleMember -function Remove-AlmDeployment
Export-ModuleMember -function Update-AlmDeployment
Export-ModuleMember -function Export-AlmDeployment
Export-ModuleMember -function ConvertTo-AlmSolutionTransformed
