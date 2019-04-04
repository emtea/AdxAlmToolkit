param (
    [string]$crmDiscoveryUrlTarget,
    [string]$usernameTarget,
    [string]$passwordTarget,
    [string]$orgNameTarget,
    [string]$filePath,
    [string]$modulePath = "../Adxstudio.Xrm.PowerShell",
    [bool]$IgnorePlugins = $true,
    # Directly import the package without extracting it
    [Switch]$NoExtract
)

# Set security protocol to TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; 

$error.clear()

# Target details
$crmDiscoveryUrl_Target = $crmDiscoveryUrlTarget
$username_Target =  $usernameTarget
$password_Target = $passwordTarget
$crmOrganisationName_Target = $orgNameTarget
$crmOrganisationDisplayName_Target = "crm target"

$targetFile = $filePath

# Install NuGet provider if not available (required for Microsoft.Xrm.Data.Powershell)
$packageProvider = Get-PackageProvider -ListAvailable -Name "NuGet" -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]"2.8.5.201" };
if(-not $packageProvider)
{
    Install-PackageProvider -Name "NuGet" -MinimumVersion "2.8.5.201" -Scope CurrentUser -Force;
}

# Install Microsoft.Xrm.Data.Powershell
$moduleInfo = Get-Module -ListAvailable -Name "Microsoft.Xrm.Data.Powershell";
if(-not $moduleInfo)
{
    Install-Module Microsoft.Xrm.Data.Powershell -Scope CurrentUser -SkipPublisherCheck -Force;
}
# Import with a prefix to prevent conflicts with ADX ALM Toolkit
Import-Module Microsoft.Xrm.Data.Powershell -Prefix Ms;

# Import ADX ALM Toolkit
Import-Module $modulePath

function Get-EnabledWebpageSdkMessageProcessingSteps
{
    param
    (
        # The CrmServiceClient object to connect to crm
        [Parameter(Mandatory=$true)]
        $conn
    )

    Write-Verbose "Retrieving steps with name *Multilanguage.WebPage* in assembly 'Adxstudio.Xrm.Plugins'";
    $steps = Get-MsCrmSdkMessageProcessingStepsForPluginAssembly -conn $conn -PluginAssemblyName "Adxstudio.Xrm.Plugins" -OnlyCustomizable;
    return $steps | Where-Object { $_.statecode -eq 'Enabled' -and $_.Name -match "MultiLanguage.WebPage" };
}

function Disable-SdkMessageProcessingStep
{
    param
    (
        # The CrmServiceClient object to connect to crm
        [Parameter(Mandatory=$true)]
        $conn,
        # The step ID
        [Parameter(Mandatory=$true)]
        $StepId
    )

    Set-MsCrmRecordState -conn $conn -EntityLogicalName sdkmessageprocessingstep -Id $StepId -StateCode Disabled -StatusCode Disabled;
}

function Enable-SdkMessageProcessingStep
{
    param
    (
        # The CrmServiceClient object to connect to crm
        [Parameter(Mandatory=$true)]
        $conn,
        # The step ID
        [Parameter(Mandatory=$true)]
        $StepId
    )

    Set-MsCrmRecordState -conn $conn -EntityLogicalName sdkmessageprocessingstep -Id $StepId -StateCode Enabled -StatusCode Enabled;
}

function importData {
    $credential = New-Object System.Management.Automation.PSCredential -ArgumentList $username_Target, ($password_Target | ConvertTo-SecureString -AsPlainText -Force)
    $crmOnlineServer = New-AlmCrmOnlineServerSetting -deploymentServiceUrl $crmDiscoveryUrl_Target -credential $credential -defaultUsername $username_Target
    $crmOrg = New-AlmOrganizationSetting -organizationName $crmOrganisationName_Target -displayName $crmOrganisationDisplayName_Target
    $almContext = New-AlmContext -name "CRM Online" -server $crmOnlineServer -organization $crmOrg
    
    if($NoExtract)
    {
        $inputPath = $targetFile;
    }
    else
    {
        # Extract the package into a folder with the same name as the package
        $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($targetFile);
        $uid = [Guid]::NewGuid().ToString().Substring(0, 6).ToLower();
        $inputPath = [System.IO.Path]::Combine($PWD, "$fileNameWithoutExtension-$uid");
        Write-Output "Extracting $targetFile into $inputPath...";
        Expand-Archive -Path $targetFile -DestinationPath $inputPath -Force
    }
    
    $crmOrganisation = Get-CrmOrganization -UniqueName $crmOrganisationName_Target -DiscoveryServiceConnection $almContext.DiscoveryServiceConnection;
    $crmOrganisationUrl = $crmOrganisation.Endpoints["WebApplication"];
    $isOnline = ($crmOrganisationUrl -match "crm(\d+)?.dynamics.com");
    # Initialise CrmServiceClient and retrieve ADX webpage plugin steps
    if($isOnline)
    {
        $conn = Connect-MsCrmOnline -Credential $credential -ServerUrl $crmOrganisationUrl;
    }
    else
    {
        $conn = Connect-MsCrmOnPremDiscovery -Credential $credential -ServerUrl $crmOrganisationUrl_Target -OrganizationName $crmOrganisationUrl;
    }
    
    $webpagePluginSteps = Get-EnabledWebpageSdkMessageProcessingSteps -conn $conn;
    $webpagePluginSteps | foreach `
    {
        Write-Output "Disabling $($_.Name)";
        Disable-SdkMessageProcessingStep -conn $conn -StepId $_.sdkmessageprocessingstepid
    }

	if($IgnorePlugins)
	{
		Import-CrmContent -OrganizationServiceConnection $almContext.OrganizationServiceConnection -InputPath $($inputPath) -UpdatesEnabled -Force -IgnorePlugins -DetailedTrace
	}
	else
	{
		Import-CrmContent -OrganizationServiceConnection $almContext.OrganizationServiceConnection -InputPath $($inputPath) -UpdatesEnabled -Force -DetailedTrace
    }

    $webpagePluginSteps | foreach `
    {
        Write-Output "Enabling $($_.Name)";
        Enable-SdkMessageProcessingStep -conn $conn -StepId $_.sdkmessageprocessingstepid
    }

    if($NoExtract)
    {
        # Nothing was created, no need to clean up
    }
    else
    {
        # Cleanup the extracted files
        Write-Output "Cleaning up extracted files at $inputPath...";
        Remove-Item -Path $inputPath -Force -Recurse
    }

    # Catch terminating errors
    trap
    {
        Write-Error "Error Occured: $_"
        exit 1
    }

    # Raise error if anything went wrong
    if ($error.Count -eq 0)
    {
        Write-Output "Script completed successfully."
    }
    else
    {
        Write-Error "Error: " $error[0].Exception.Message
        Write-Error "Error: " $error[0].Exception.ToString()
        exit 1
    }
}

importData