param (
    [string]$crmDiscoveryUrlSource,
    [string]$usernameSource,
    [string]$passwordSource,
    [string]$orgNameSource,
    [string]$filePath,
    [string]$xmlPath,
    [string]$modulePath = "../Adxstudio.Xrm.PowerShell"
)

# Set security protocol to TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; 

$error.clear()

# Source detials
$crmDiscoveryUrl_Source = $crmDiscoveryUrlSource
$username_Source =  $usernameSource
$password_Source = $passwordSource
$crmOrganisationName_Source = $orgNameSource
$crmOrganisationDisplayName_Source = "crm Source"

$targetFile = $filePath

$dataExportFetch = [IO.File]::ReadAllText($xmlPath)

Import-Module $modulePath 

function exportData {
    $credential = New-Object System.Management.Automation.PSCredential -ArgumentList $username_Source, ($password_Source | ConvertTo-SecureString -AsPlainText -Force)
    $crmOnlineServer = New-AlmCrmOnlineServerSetting -deploymentServiceUrl $crmDiscoveryUrl_Source -credential $credential -defaultUsername $username_Source
    $crmOrg = New-AlmOrganizationSetting -organizationName $crmOrganisationName_Source -displayName $crmOrganisationDisplayName_Source
    $almContext = New-AlmContext -name "CRM Online" -server $crmOnlineServer -organization $crmOrg

    Export-CrmContent -OutputPath $($targetFile) -OrganizationServiceConnection $almContext.OrganizationServiceConnection -ContentFetchXml $dataExportFetch -ExcludeMetadata 

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

exportData
