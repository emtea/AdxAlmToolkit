Function Test-Any()
{
	begin {
		$any = $false
	}
	process {
		$any = $true
	}
	end {
		$any
	}
}

Function Invoke-SqlQuery
{
	<#
	.SYNOPSIS
	Invokes a SQL query by text. Requires Microsoft® Windows PowerShell Extensions for Microsoft® SQL Server® 2012.

	.PARAMETER SqlCredential
	SQL authentication based user credentials.

	.EXAMPLE
	PS C:\> Invoke-SqlQuery -SqlServerName "crm2011" -Query "sp_helpdb 'MSCRM_CONFIG'"
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		$sqlServerName,

		[Parameter(Position=1)]
		$sqlCredential,
		
		[Parameter(Mandatory=$true, Position=2)]
		$query,

		[Parameter()]
		$database,

		[Parameter()]
		$queryTimeout
	)

	$baseParams = @{
		"Query" = $query;
		"ServerInstance" = $sqlServerName;
	}

	if ($queryTimeout)
	{
		$baseParams.Add("QueryTimeout", $queryTimeout)
	}

	if ($sqlCredential)
	{
		$baseParams.Add("Username", $sqlCredential.GetNetworkCredential().UserName)
		$baseParams.Add("Password", $sqlCredential.GetNetworkCredential().Password)
	}

	if ($database)
	{
		$baseParams.Add("Database", $database)
	}

	Invoke-Sqlcmd @baseParams
}

Function Test-CrmIsDeploymentAdmin
{
	<#
	.SYNOPSIS
	Checks that a user is a member of deployment administrators.
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[Alias("Connection")]
		$deploymentServiceConnection
	)

	$server = Get-CrmDeploymentEntity -Connection $deploymentServiceConnection -DeploymentEntityType Server

	Write-Output $server
}

Function Get-CrmConnectionString
{
	<#
	.SYNOPSIS
	Converts a PSCrmConnection object to a connection-string. The connection must be an organization service connection.
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		$connection
	)

	if ($connection.Connection)
	{
		$serviceUri = $connection.Connection.ServiceUri
		$timeout = $connection.Connection.Timeout
		$config = Get-CrmOrganizationConfig -OrganizationServiceUrl $serviceUri

		if ($config.AuthenticationType -eq "ActiveDirectory")
		{
			$wc = $connection.Connection.ClientCredentials.Windows.ClientCredential

			Write-Output ("Url={0}; Domain={1}; Username={2}; Password={3}; Timeout={4};" -f $serviceUri, $wc.Domain, $wc.UserName, $wc.Password, $timeout)
		}
		else
		{
			$uc = $connection.Connection.ClientCredentials.UserName

			Write-Output ("Url={0}; Username={1}; Password={2}; Timeout={3};" -f $serviceUri, $uc.UserName, $uc.Password, $timeout)
		}
	}
	else
	{
		$serviceUri = $connection.ServiceUri
		$timeout = $connection.Timeout
		$nc = $connection.Credential.GetNetworkCredential()
		$config = Get-CrmOrganizationConfig -OrganizationServiceUrl $serviceUri

		if ($config.AuthenticationType -eq "ActiveDirectory")
		{
			Write-Output ("Url={0}; Domain={1}; Username={2}; Password={3}; Timeout={4};" -f $serviceUri, $nc.Domain, $nc.UserName, $nc.Password, $timeout)
		}
		else
		{
			Write-Output ("Url={0}; Username={1}; Password={2}; Timeout={3};" -f $serviceUri, $connection.Credential.UserName, $nc.Password, $timeout)
		}
	}
}

Function Restore-CrmSqlDatabase
{
	<#
	.SYNOPSIS
	Restores an organization backup file to the database server.

	.PARAMETER SqlBackupFile
	The SQL .bak file exported from an existing organization. This file must be located on the SQL server machine.

	.PARAMETER SqlDataPath
	The target folder of the restored database data files on the SQL server machine.

	.EXAMPLE
	PS C:\> Import-Module SQLPS
	PS C:\> $sqlDatabaseName = Restore-CrmSqlDatabase -SqlServerName "crm2011" -OrganizationName "Contoso" -SqlBackupFile "C:\backups\AdventureWorks_MSCRM.bak" -SqlDataPath "C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\"
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true)]
		$sqlServerName = "crm2011",

		[Parameter()]
		$sqlCredential,

		[Parameter(Mandatory=$true)]
		[Alias("OrgName")]
		$organizationName,

		[Parameter(Mandatory=$true)]
		$sqlBackupFile,

		[Parameter(Mandatory=$true)]
		$sqlDataPath = "C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\"
	)

	# check if the database already exists

	$filteredToUnderscore = $organizationName -replace "[\-]", "_"
	$sqlDatabaseName = "${filteredToUnderscore}_MSCRM"
	$dbExistsQuery = "select Id = db_id('$sqlDatabaseName')"

	Write-Verbose "Invoke: $dbExistsQuery"

	$dbExists = Invoke-SqlQuery $sqlServerName $sqlCredential $dbExistsQuery

	if ($dbExists.Id -ne [System.DBNull]::Value)
	{
		throw New-Object System.ApplicationException("The database already exists.")
	}

	$restoreQuery = @"
RESTORE DATABASE [$sqlDatabaseName]
  FROM DISK = N'$sqlBackupFile' WITH FILE = 1,
  MOVE N'mscrm' TO N'${sqlDataPath}${sqlDatabaseName}.mdf',
  MOVE N'mscrm_log' TO N'${sqlDataPath}${sqlDatabaseName}.ldf',
  NOUNLOAD,  STATS = 10
"@

	Write-Verbose "Invoke: $restoreQuery"

	Invoke-SqlQuery $sqlServerName $sqlCredential $restoreQuery

	return $sqlDatabaseName
}

Function Convert-CrmOrganizationUniqueNameToWebApplicationUrl
{
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[Alias("Connection")]
		$discoveryServiceConnection,

		[Parameter(Mandatory=$true, Position=1)]
		$uniqueName
	)

	$request = New-Object Microsoft.Xrm.Sdk.Discovery.RetrieveOrganizationRequest
	$request.UniqueName = $uniqueName
	$response = Invoke-CrmDiscoveryService -Connection $discoveryServiceConnection -Request $request
	$url = $response.Detail.Endpoints[[Microsoft.Xrm.Sdk.Discovery.EndpointType]::WebApplication]

	Write-Output $url
}

Function Get-CrmAnyErrorRetryPolicy
{
	[CmdletBinding()]
	param (
		[int]
		$retryCount = 10,

		[System.TimeSpan]
		$retryInterval = [System.TimeSpan]::Parse("00:00:30")
	)

	$referencedAssemblies = Join-Path $PSScriptRoot "Microsoft.Practices.TransientFaultHandling.Core.dll" -Resolve

	$source = @"
public class AnyErrorDetectionStrategy : Microsoft.Practices.TransientFaultHandling.ITransientErrorDetectionStrategy
{
	public virtual bool IsTransient(System.Exception ex) { return true; }
}
"@

	Add-Type -Path $referencedAssemblies
	Add-Type -TypeDefinition $source -ReferencedAssemblies $referencedAssemblies

	$strategy = New-Object AnyErrorDetectionStrategy
	$fixedInterval = New-Object Microsoft.Practices.TransientFaultHandling.FixedInterval($null, $retryCount, $retryInterval, $true)
	return New-Object Microsoft.Practices.TransientFaultHandling.RetryPolicy($strategy, $fixedInterval)
}

Function New-CrmDeployment
{
	<#
	.SYNOPSIS
	Creates a new CRM organization.

	.EXAMPLE
	PS C:\> $connection = Get-CrmConnection "Url=...; Domain=...; Username=...; Password=;"
	PS C:\> New-CrmDeployment -DeploymentServiceConnection $connection -OrganizationName "AdventureWorks" -SqlServerName "crm2011" -SrsUrl" http://crm2011/reportserver"

	.EXAMPLE
	PS C:\> $connection = Get-CrmConnection "Url=...; Domain=...; Username=...; Password=;"
	PS C:\> $currency = @{ "BaseCurrencyCode" = "EUR"; "BaseCurrencyName" = "euro"; "BaseCurrencyPrecision" = 2; "BaseCurrencySymbol" = "€" } # Note: use PowerShell ISE for unicode support
	PS C:\> New-CrmDeployment -DeploymentServiceConnection $connection -OrganizationName "AdventureWorks" -SqlServerName "crm2011" -SrsUrl" http://crm2011/reportserver" -Currency $currency
	#>

	[CmdletBinding(DefaultParameterSetName="ByDeployment")]
	Param (
		[Parameter(Mandatory=$true, Position=0, ParameterSetName="ByDeployment")]
		[Alias("Connection")]
		$deploymentServiceConnection,
		
		[Parameter(Mandatory=$true, Position=1, ParameterSetName="ByDeployment")]
		[Alias("OrgName")]
		$organizationName,

		[Parameter(Mandatory=$true, Position=2, ParameterSetName="ByDeployment")]
		$displayName,
		
		[Parameter(Mandatory=$false, Position=3)]
		$sqlServerName = "crm2011",
		
		[Parameter(Mandatory=$false, Position=4)]
		$srsUrl = "http://crm2011/reportserver",

		[Parameter(Mandatory=$false)]
		$currency,

		[Parameter(Mandatory=$false)]
		[int]
		$baseLanguageCode,

		[Parameter(Mandatory=$false)]
		[int[]]
		$languageCode,

		[Parameter(Mandatory=$false)]
		$sqlDatabaseName,

		[Parameter(Mandatory=$false)]
		[string[]]
		$customizationFile,
		
		[Parameter(Mandatory=$false)]
		[string[]]
		$dataFile,
		
		[Parameter(Mandatory=$false)]
		$adminFirstName = "Domain",
		
		[Parameter(Mandatory=$false)]
		$adminLastname = "Administrator",

		[Parameter(Mandatory=$true, ParameterSetName="ByOrganization")]
		$organizationServiceConnection,
		
		[Parameter(Mandatory=$false)]
		[System.Collections.Hashtable[]]
		$newUsers,

		[Parameter(Mandatory=$false)]
		$dataEncryptionKey,

		[Parameter()]
		[Switch]
		$updatesEnabled,

		[Parameter()]
		[Switch]
		$importSolutionAsJob,

		[Parameter()]
		$importSolutionTimeout = "00:30:00",

		[Parameter()]
		$preImportStatePath,

		[Parameter()]
		$crmDataCopyToolPath
	)

	# retry 10 times with a 30 sec pause interval

	$retrieveOrgRetryPolicy = New-AlmRetryPolicy 10 "00:00:30" $false -MessageToRetry "Microsoft Dynamics CRM has experienced an error\."

	if ($organizationServiceConnection)
	{
		# use an existing organization

		if ($organizationServiceConnection.Connection)
		{
			$orgUrl = $organizationServiceConnection.Connection.ServiceUri
		}
		else
		{
			$orgUrl = $organizationServiceConnection.ServiceUri
		}
	}
	else
	{
		# check if the organiztion already exists
	
		$existingOrgs = Get-CrmDeploymentEntity -Connection $deploymentServiceConnection -DeploymentEntityType Organization
		$orgExists = $existingOrgs | Where-Object { $_.Name -eq $organizationName } | Test-Any
	
		if ($orgExists)
		{
			throw New-Object System.ApplicationException("An organization with the name '$organizationName' already exists.")
		}
	
		$baseParams = @{
			"Connection" = $deploymentServiceConnection;
			"UniqueName" = $organizationName;
			"DisplayName" = $displayName;
			"SqlServerName" = $sqlServerName;
			"SrsUrl" = $srsUrl;
			"ReturnAsUniqueName" = $true
		}

		if ($sqlDatabaseName)
		{
			# import a new organization from an existing database

			Write-Host "Importing organization..."

			$baseParams.Add("DatabaseName", $sqlDatabaseName);

			$orgUniqueName = Import-CrmOrganization @baseParams
		}
		else
		{
			# create a new organization

			Write-Host "Creating organization..."

			$baseParams.Add("BaseLanguageCode", $baseLanguageCode)

			if ($currency)
			{
				$currency.GetEnumerator() | ForEach-Object { $baseParams.Add($_.Name, $_.Value) }
			}

			$orgUniqueName = New-CrmOrganization @baseParams
		}

		$retrieveOrgUrlAction = {
			Write-Host "Retrieving organization URL..."
			Convert-CrmOrganizationUniqueNameToWebApplicationUrl $deploymentServiceConnection $orgUniqueName
		}

		$orgUrl = Invoke-AlmRetryPolicy -RetryPolicy $retrieveOrgRetryPolicy $retrieveOrgUrlAction

		if (-not $orgUrl)
		{
			throw New-Object System.ApplicationException("Failed to create or retrieve the organization.")
		}

		Write-Host $orgUrl

		# convert the deployment connection to an organization connection

		if ($deploymentServiceConnection.Connection)
		{
			$organizationServiceConnection = Get-CrmConnection -Connection (New-Object Microsoft.Xrm.Client.CrmConnection)
			$organizationServiceConnection.Connection.ServiceUri = $orgUrl
			$organizationServiceConnection.Connection.ClientCredentials = $deploymentServiceConnection.Connection.ClientCredentials
			$organizationServiceConnection.Connection.DeviceCredentials = $deploymentServiceConnection.Connection.DeviceCredentials
			if ($deploymentServiceConnection.Connection.Timeout) { $organizationServiceConnection.Connection.Timeout = $deploymentServiceConnection.Connection.Timeout }
			if ($deploymentServiceConnection.Connection.CallerId) { $organizationServiceConnection.Connection.CallerId = $deploymentServiceConnection.Connection.CallerId }
		}
		else
		{
			$organizationServiceConnection = Get-CrmConnection -Url $orgUrl -Credential $deploymentServiceConnection.Credential -Timeout $deploymentServiceConnection.Timeout
		}
	}

	# disable error reporting for the organization

	$retrieveOrgAction = {
		Write-Host "Retrieving organization..."
		Get-CrmEntity -Connection $organizationServiceConnection -EntityLogicalName organization | Select-Object -First 1
	}

	$org = Invoke-AlmRetryPolicy -RetryPolicy $retrieveOrgRetryPolicy $retrieveOrgAction

	if ($org)
	{
		Write-Host "Updating organization..."

		$reportScriptErrors = New-Object Microsoft.Xrm.Sdk.OptionSetValue -ArgumentList 3

		$newOrg = New-Object Microsoft.Xrm.Sdk.Entity -ArgumentList "organization"
		$newOrg.Id = $org.Id
		$newOrg.Attributes["organizationid"] = $org.Id
		$newOrg.Attributes["reportscripterrors"] = [Microsoft.Xrm.Sdk.OptionSetValue] $reportScriptErrors
		$newOrg | Format-CrmEntity

		$newOrgResponse = Set-CrmEntity -Connection $organizationServiceConnection -Entity $newOrg
	}

	# provision languages by locale code

	if ($languageCode)
	{
		$languageCode | ForEach-Object {
			Write-Host ("Provisioning language: {0}" -f $_)
			Invoke-CrmOrganizationService -Connection $organizationServiceConnection -RequestName "ProvisionLanguage" -Parameters @{ "Language" = [int] $_ }
		}
	}

	# fix the administrator name

	Write-Host "Updating administrator..."

	$userId = Get-CrmWhoAmI -Connection $organizationServiceConnection -AttributeName "UserId"

	if ($userId)
	{
		$newUser = New-Object Microsoft.Xrm.Sdk.Entity -ArgumentList "systemuser"
		$newUser.Id = $userId
		$newUser.Attributes["systemuserid"] = $userId
		$newUser.Attributes["firstname"] = $adminFirstName
		$newUser.Attributes["lastname"] = $adminLastname
		$newUser | Format-CrmEntity

		$newUserResponse = Set-CrmEntity -Connection $organizationServiceConnection -Entity $newUser
	}

	if ($dataEncryptionKey)
	{
		Write-Host "Setting data encryption key..."

		Invoke-CrmOrganizationService -Connection $organizationServiceConnection -RequestName "SetDataEncryptionKey" -Parameters @{ ChangeEncryptionKey = $false; EncryptionKey = $dataEncryptionKey }
	}

	# create custom systemusers
	
	if ($newUsers)
	{
		Write-Host "Creating custom systemuser..."
		
		$existingUsers = Get-CrmEntity -Connection $organizationServiceConnection -EntityLogicalName "systemuser" -ColumnSet "domainname", "internalemailaddress", "fullname"
		$existingUsers | Format-CrmEntity

		$businessUnitId = Get-CrmWhoAmI -Connection $organizationServiceConnection -AttributeName "BusinessUnitId" -EntityLogicalname "businessunit"

		$newUsers | ForEach-Object { $_.Set_Item("businessunitid", $businessUnitId) }
		$newUserEntities = $newUsers | Select-CrmEntity "systemuser" | Where-Object { $u = $_; -not ($existingUsers | Where-Object { $u.Attributes["domainname"] -eq $_.Attributes["domainname"] } | Test-Any) }
		
		if ($newUserEntities)
		{
			$newUserEntities | Format-CrmEntity
		
			$newUserIds = $newUserEntities | New-CrmEntity -Connection $organizationServiceConnection

			Write-Host $newUserIds

			$role = Get-CrmEntity -Connection $organizationServiceConnection -EntityLogicalName "role" | Where-Object { $_.Attributes["name"] -eq "System Administrator" } | Select-Object -First 1

			if ($role)
			{
				# Associate role to a system user

				foreach ($userId in $newUserIds)
				{
					Write-Host "Associating System Administrator Role to systemuser with ID ${userId}"

					$userEntityReference = New-Object Microsoft.Xrm.Sdk.EntityReference
					$userEntityReference.LogicalName = "systemuser"
					$userEntityReference.Id = $userId

					$relationship = New-Object Microsoft.Xrm.Sdk.Relationship
					$relationship.SchemaName = "systemuserroles_association"

					$roleEntityReference = New-Object Microsoft.Xrm.Sdk.EntityReference
					$roleEntityReference.LogicalName = $role.EntityLogicalName
					$roleEntityReference.Id = $role.Id

					$relatedEntities = New-Object Microsoft.Xrm.Sdk.EntityReferenceCollection
					$relatedEntities.Add($roleEntityReference)

					$associateRequest = New-Object Microsoft.Xrm.Sdk.Messages.AssociateRequest
					$associateRequest.Target = $userEntityReference
					$associateRequest.Relationship = $relationship
					$associateRequest.RelatedEntities = $relatedEntities

					$associateResponse = Invoke-CrmOrganizationService -Connection $organizationServiceConnection -Request $associateRequest
				}
			}
		}
	}

	if ($customizationFile)
	{
		Write-Host "Importing solutions..."

		if ($importSolutionAsJob)
		{
			$importJob = Import-CrmSolutionAsJob $organizationServiceConnection $customizationFile -ImportSolutionTimeout $importSolutionTimeout
		}
		else
		{
			$importJob = Import-CrmSolutionSynchronous $organizationServiceConnection $customizationFile
		}

		Write-Host "Imported solutions:"
		$importJob | ft | Out-Host
	}

	if ($dataFile)
	{
		Write-Host "Importing data..."

		if ($crmDataCopyToolPath)
		{
			$dataFile | % {
				# execute import by executable
				$connectionString = Get-CrmConnectionString $organizationServiceConnection
				if ($updatesEnabled) { $updatesEnabledParam = "/updatesEnabled" }

				Write-Host "Calling: $crmDataCopyToolPath /action:import /connectionString:... /in:$_ /preImportStatePath:$preImportStatePath /force $updatesEnabledParam"

				$params = ("/action:import", "/connectionString:$connectionString", "/in:$_", "/preImportStatePath:$preImportStatePath", "/force", $updatesEnabledParam)

				& $crmDataCopyToolPath $params | Out-Host
			}
		}
		else
		{
			$dataFile | % {
				# execute import by cmdlet
				Import-CrmContent -Connection $organizationServiceConnection -InputPath $_ -PreImportStatePath $preImportStatePath -UpdatesEnabled:$updatesEnabled -Force -Serial
			}
		}
	}

	if ($customizationFile)
	{
		Write-Host "Publishing all customizations..."

		$publish = Publish-CrmCustomization -Connection $organizationServiceConnection
	}

	# return the organization URL

	Write-Output $orgUrl
}

Function Import-CrmSolutionSynchronous
{
	<#
	.SYNOPSIS
	Imports a solution by waiting for the import service call to complete.

	.EXAMPLE
	PS C:\> Import-CrmSolutionSynchronous -OrganizationServiceConnection $connection -CustomizationPath $customizationFile
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[Alias("Connection")]
		$organizationServiceConnection,

		[Parameter(Mandatory=$false, Position=1)]
		[string[]]
		$customizationFile
	)

	$importJob = Import-CrmSolution -Connection $organizationServiceConnection -CustomizationPath $customizationFile -PublishWorkflows

	if (-not $importJob)
	{
		throw New-Object System.ApplicationException("Failed to import solution(s).")
	}
	else
	{
		$importJob | Watch-CrmImportJob -Connection $organizationServiceConnection -ImportTimeout "00:01:00" | % {
			$job = $_

			if ($job.Data)
			{
				$importError = $job.Data | Get-CrmImportJobError

				if ($importError)
				{
					throw New-Object System.ApplicationException($importError)
				}
			}

			if ($job.State -eq "Completed")
			{
				# pass on only the completed job

				Write-Output $job
			}
		}
	}
}

Function Import-CrmSolutionAsJob
{
	<#
	.SYNOPSIS
	Imports a solution using a background job and monitors its progress.

	.EXAMPLE
	PS C:\> Import-CrmSolutionAsJob -OrganizationServiceConnection $connection -CustomizationPath $customizationFile
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[Alias("Connection")]
		$organizationServiceConnection,

		[Parameter(Mandatory=$false, Position=1)]
		[string[]]
		$customizationFile,

		$progressInterval = "00:00:03",

		$importSolutionTimeout = "00:30:00"
	)

	$activity = "Importing Solution"

	Import-CrmSolution -Connection $organizationServiceConnection -CustomizationPath $customizationFile -PublishWorkflows -AsJob | Watch-CrmImportJob -Connection $organizationServiceConnection -ProgressInterval $progressInterval -ImportTimeout $importSolutionTimeout | % {
		$job = $_

		if ($job.Data)
		{
			$importError = $job.Data | Get-CrmImportJobError

			if ($importError)
			{
				Write-Progress -Id 3 -Activity $activity -Completed
				Write-Progress -Id 2 -Activity $activity -Completed
				throw New-Object System.ApplicationException($importError)
			}
		}

		if (($job.State -eq "Failed") -or ($job.State -eq "Blocked"))
		{
			Write-Progress -Id 3 -Activity $activity -Completed
			Write-Progress -Id 2 -Activity $activity -Completed
			throw ("Failed to import solution: {0}" -f $job.SolutionName)
		}

		if ($job.SolutionName) { $status = $job.SolutionName } else { $status = $job.State }
		Write-Progress -Id 3 -Activity "Importing Solution" -Status $status -percentComplete $job.Progress

		if ($job.State -eq "Completed")
		{
			# pass on only the completed job

			Write-Output $job
		}
	}

	Write-Progress -Id 3 -Activity $activity -Completed
}

Function Remove-CrmDeployment
{
	<#
	.SYNOPSIS
	Removes an existing CRM organization.

	.PARAMETER SqlCredential
	SQL authentication based user credentials of a user with access to drop the CRM database.

	.EXAMPLE
	PS C:\> $connection = Get-CrmConnection "Url=...; Domain=...; Username=...; Password=;"
	PS C:\> $organizationid = Get-CrmEntityInstanceId "AdventureWorks02"
	PS C:\> Remove-CrmDeployment -DeploymentServiceConnection $connection -EntityInstanceId $organizationid
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[Alias("Connection")]
		$deploymentServiceConnection,

		[Parameter(Mandatory=$true)]
		$entityInstanceId,

		[Parameter()]
		$sqlCredential,

		[Parameter()]
		[Switch]
		$force
	)

	# display organization details

	try
	{
		$orgResults = Get-CrmDeploymentEntity -Connection $deploymentServiceConnection -DeploymentEntityType Organization -EntityInstanceId $entityInstanceId
	}
	catch [System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.Deployment.DeploymentServiceFault]]]
	{
		Write-Verbose $_.Exception.Message
	}

	if ($orgResults)
	{
		$orgResults
	
		if ($force)
		{
			$result = 0
		}
		else
		{
			$title = "Delete organization"
			$message = "Do you want to delete the organization?"

			$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Continue"
			$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel"
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

			$result = $host.ui.PromptForChoice($title, $message, $options, 1) 
		}

		switch ($result)
		{
			0 { Remove-CrmDeploymentInternal -Connection $deploymentServiceConnection -Organization $orgResults -SqlCredential $sqlCredential -Force:$force }
			1 { "Cancelled." }
		}
	}
}

Function Remove-CrmDeploymentInternal
{
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[Alias("Connection")]
		$deploymentServiceConnection,
		
		[Parameter(Mandatory=$true)]
		$organization,

		[Parameter()]
		$sqlCredential,

		[Parameter()]
		[Switch]
		$force
	)

	# load SQL values from the organization deployment entity

	$entityInstanceId = Get-CrmEntityInstanceId $organization.UniqueName
	$sqlServerName = $organization.SqlServerName
	$databaseName = $organization.DatabaseName

	$removedOrg = Remove-CrmOrganization -Connection $deploymentServiceConnection -EntityInstanceId $entityInstanceId
	
	if ($removedOrg)
	{
		if (Get-Command "Invoke-Sqlcmd" -errorAction SilentlyContinue)
		{
			if ($sqlServerName -and $databaseName)
			{
				Remove-SqlDatabase $sqlServerName $databaseName -SqlCredential $sqlCredential -Force:$force
			}
		}
		else
		{
			Write-Host "For automated database deletion, install the Microsoft® Windows PowerShell Extensions for Microsoft® SQL Server® 2012."
			Write-Host "The '$databaseName' database on the '$sqlServerName' server can now be dropped manually."
		}

		Write-Output $removedOrg
	}
}

Function Remove-SqlDatabase
{
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$sqlServerName,
		
		[Parameter(Mandatory=$true,Position=1)]
		$databaseName,

		[Parameter(Position=2)]
		$sqlCredential,

		[Parameter()]
		[Switch]
		$force
	)
	
	# display database details
	
	$helpQuery = "sp_helpdb '$databaseName'"

	$dbResults = Invoke-SqlQuery $sqlServerName $sqlCredential $helpQuery

	if ($dbResults)
	{
		$dbResults
	
		if ($force)
		{
			$result = 0
		}
		else
		{
			$title = "Delete database"
			$message = "Do you want to delete the '$databaseName' database on the '$sqlServerName' server?"

			$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Continue"
			$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel"
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

			$result = $host.ui.PromptForChoice($title, $message, $options, 1) 
		}

		switch ($result)
		{
			0 { Remove-SqlDatabaseInternal $sqlServerName $databaseName -SqlCredential $sqlCredential }
			1 { "Cancelled." }
		}
	}
}

Function Remove-SqlDatabaseInternal
{
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$sqlServerName,
		
		[Parameter(Mandatory=$true,Position=1)]
		$databaseName,

		[Parameter(Position=2)]
		$sqlCredential
	)
	
	# delete database history
	
	$deleteHistoryDbQuery = "exec msdb.dbo.sp_delete_database_backuphistory @database_name = N'$databaseName'"
	Write-Verbose "Deleting database history"
	Write-Verbose "Invoke: $deleteHistoryDbQuery"
	
	Invoke-SqlQuery $sqlServerName $sqlCredential $deleteHistoryDbQuery

	# drop connections to the database
	
	$alterDbQuery = "alter database [$databaseName] set single_user with rollback immediate"
	Write-Verbose "Dropping database connections"
	Write-Verbose "Invoke: $alterDbQuery"
	
	Invoke-SqlQuery $sqlServerName $sqlCredential $alterDbQuery

	# delete the database
	
	$dropDbQuery = "drop database [$databaseName]"
	Write-Verbose "Deleting database"
	Write-Verbose "Invoke: $dropDbQuery"
	
	Invoke-SqlQuery $sqlServerName $sqlCredential $dropDbQuery
}

Function Remove-CrmContent
{
	<#
	.SYNOPSIS
	Removes content from a target organization that does not exist in the specified source data.

	.EXAMPLE
	PS C:\> $connection = Get-CrmConnection "Url=...; Domain=...; Username=...; Password=;"
	PS C:\> $dataFile = "C:\Data\MyOrgData.zip"
	PS C:\> Remove-CrmContent -OrganizationServiceConnection $connection -DataFile $dataFile -Exclude ("customeraddress", "connection")

	.EXAMPLE
	PS C:\> $connection = Get-CrmConnection "Url=...; Domain=...; Username=...; Password=;"
	PS C:\> $dataFile = "C:\Data\MyOrgData.zip"
	PS C:\> Remove-CrmContent -OrganizationServiceConnection $connection -DataFile $dataFile -Include ("account", "contact")
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[Alias("Connection")]
		$organizationServiceConnection,

		[Parameter(Mandatory=$true, Position=1)]
		$dataFile,

		[Parameter()]
		$exclude,

		[Parameter()]
		$include,

		[Parameter()]
		$pageCount,

		[Parameter()]
		[Switch]
		$includeUnmanagedSolutionContent,

		[Parameter()]
		[Switch]
		$force
	)

	# retrieve the entity Ids from the data file

	$source = Read-CrmContent (Resolve-Path $dataFile)
	$sourceIds = $source | ForEach-Object { $_.Entity.ToEntityReference() }

	# retrieve the exportable entities from the target organization

	$target = Get-CrmContent -Connection $organizationServiceConnection -Generalized -Exclude $exclude -Include $include -PageCount $pageCount -IncludeUnmanagedSolutionContent:$includeUnmanagedSolutionContent

	# find and display the difference

	$targetsNotInSource = $target | Where-Object { -not $sourceIds.Contains($_.ToEntityReference()) }
	$count = $targetsNotInSource.Length

	Write-Host
	Write-Host "Found $count entities to delete..."
	Write-Host

	if ($count -gt 0)
	{
		$targetsNotInSource | Format-CrmEntity

		if ($force)
		{
			$result = 0
		}
		else
		{
			$title = "Delete content"
			$message = "Do you want to delete the $count entities?"

			$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Continue"
			$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel"
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

			$result = $host.ui.PromptForChoice($title, $message, $options, 1) 
		}

		switch ($result)
		{
			0 { $targetsNotInSource | ForEach-Object { $_.ToEntityReference() } | Remove-CrmEntity -Connection $organizationServiceConnection }
			1 { "Cancelled." }
		}
	}
}

Export-ModuleMember -function Invoke-SqlQuery
Export-ModuleMember -function Test-CrmIsDeploymentAdmin
Export-ModuleMember -function Restore-CrmSqlDatabase
Export-ModuleMember -function Get-CrmAnyErrorRetryPolicy
Export-ModuleMember -function New-CrmDeployment
Export-ModuleMember -function Import-CrmSolutionSynchronous
Export-ModuleMember -function Import-CrmSolutionAsJob
Export-ModuleMember -function Remove-CrmDeployment
Export-ModuleMember -function Remove-SqlDatabase
Export-ModuleMember -function Remove-CrmContent
