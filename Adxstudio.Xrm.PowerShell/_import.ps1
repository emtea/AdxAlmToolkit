param (
    [string]$crmDiscoveryUrlTarget,
    [string]$usernameTarget,
    [string]$passwordTarget,
    [string]$orgNameTarget,
    [string]$filePath
    #[string]$modulePath
)

$error.clear()

# Target details
$crmDiscoveryUrl_Target = $crmDiscoveryUrlTarget
$username_Target =  $usernameTarget
$password_Target = $passwordTarget
$crmOrganisationName_Target = $orgNameTarget
$crmOrganisationDisplayName_Target = "crm target"

$targetFile = $filePath

Import-Module "C:\Program Files (x86)\Adxstudio\ALM Toolkit\1.0.0017\Adxstudio.Xrm.PowerShell"

function importData {
    $credential = New-Object System.Management.Automation.PSCredential -ArgumentList $username_Target, ($password_Target | ConvertTo-SecureString -AsPlainText -Force)
    $crmOnlineServer = New-AlmCrmOnlineServerSetting -deploymentServiceUrl $crmDiscoveryUrl_Target -credential $credential -defaultUsername $username_Target
    $crmOrg = New-AlmOrganizationSetting -organizationName $crmOrganisationName_Target -displayName $crmOrganisationDisplayName_Target
    $almContext = New-AlmContext -name "CRM Online" -server $crmOnlineServer -organization $crmOrg

    Import-CrmContent -OrganizationServiceConnection $almContext.OrganizationServiceConnection -InputPath $($targetFile) -UpdatesEnabled -Force -IgnorePlugins -DetailedTrace

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
