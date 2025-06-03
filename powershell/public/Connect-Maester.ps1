function Connect-Maester {
    <#
    .SYNOPSIS
    Enhanced helper method to connect to Microsoft Graph using Connect-MgGraph with support for registered applications.

    .DESCRIPTION
    Use this cmdlet to connect to Microsoft Graph using Connect-MgGraph with various authentication methods including
    registered Azure applications with client secrets or certificates.

    This command is completely optional if you are already connected to Microsoft Graph and other services using Connect-MgGraph with the required scopes.

    .EXAMPLE
    Connect-Maester

    Connects to Microsoft Graph using Connect-MgGraph with the required scopes (interactive authentication).

    .EXAMPLE
    Connect-Maester -ClientId "12345678-1234-1234-1234-123456789012" -TenantId "87654321-4321-4321-4321-210987654321" -ClientSecret "your-client-secret"

    Connects using a registered application with client secret.

    .EXAMPLE
    Connect-Maester -ClientId "12345678-1234-1234-1234-123456789012" -TenantId "87654321-4321-4321-4321-210987654321" -CertificateThumbprint "ABC123..."

    Connects using a registered application with certificate authentication.

    .EXAMPLE
    Connect-Maester -ClientId "12345678-1234-1234-1234-123456789012" -TenantId "87654321-4321-4321-4321-210987654321" -CertificatePath "C:\path\to\cert.pfx" -CertificatePassword "password"

    Connects using a registered application with certificate file and password.

    .EXAMPLE
    Connect-Maester -UseDeviceCode

    Connects using the device code flow.

    .EXAMPLE
    Connect-Maester -UseManagedIdentity

    Connects using Azure Managed Identity (for Azure Automation, Azure VMs, etc.).

    .LINK
    https://maester.dev/docs/commands/Connect-Maester
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Colors are beautiful')]
    [Alias('Connect-MtGraph', 'Connect-MtMaester')]
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    param(
        # Specifies the services to connect to. Default is Graph only.
        [ValidateSet('Graph', 'Azure', 'ExchangeOnline', 'All')]
        [string[]] $Service = @('Graph'),

        # If specified, the cmdlet will include the scope to send email (Mail.Send).
        [switch] $SendMail,

        # If specified, the cmdlet will include the scope to send a channel message in Teams (ChannelMessage.Send).
        [switch] $SendTeamsMessage,

        # If specified, the cmdlet will include the scopes for read write API endpoints.
        [switch] $Privileged,

        # If specified, the cmdlet will use the device code flow to authenticate to Graph and Azure.
        [Parameter(ParameterSetName = 'DeviceCode')]
        [switch] $UseDeviceCode,

        # If specified, the cmdlet will use Managed Identity to authenticate.
        [Parameter(ParameterSetName = 'ManagedIdentity')]
        [switch] $UseManagedIdentity,

        # The Application (Client) ID of the registered Azure application.
        [Parameter(ParameterSetName = 'ClientSecret', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Certificate', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CertificateFile', Mandatory = $true)]
        [string] $ClientId,

        # The Directory (Tenant) ID of the Azure AD tenant.
        [Parameter(ParameterSetName = 'ClientSecret', Mandatory = $true)]
        [Parameter(ParameterSetName = 'Certificate', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CertificateFile', Mandatory = $true)]
        [string] $TenantId,

        # The client secret of the registered Azure application.
        [Parameter(ParameterSetName = 'ClientSecret', Mandatory = $true)]
        [SecureString] $ClientSecret,

        # The thumbprint of the certificate used for authentication.
        [Parameter(ParameterSetName = 'Certificate', Mandatory = $true)]
        [string] $CertificateThumbprint,

        # The path to the certificate file (.pfx).
        [Parameter(ParameterSetName = 'CertificateFile', Mandatory = $true)]
        [string] $CertificatePath,

        # The password for the certificate file.
        [Parameter(ParameterSetName = 'CertificateFile')]
        [SecureString] $CertificatePassword,

        # The national cloud environment to connect to.
        [ValidateSet('Global', 'China', 'Germany', 'USGov', 'USGovDOD')]
        [string] $Environment = 'Global',

        # The Azure cloud environment for Azure PowerShell connections.
        [ValidateSet('AzureCloud', 'AzureChinaCloud', 'AzureGermanCloud', 'AzureUSGovernment')]
        [string] $AzureEnvironment = 'AzureCloud',

        # The Exchange Online environment name.
        [ValidateSet('O365Default', 'O365GermanyCloud', 'O365China', 'O365USGovGCCHigh', 'O365USGovDoD')]
        [string] $ExchangeEnvironmentName = 'O365Default'
    )

    # Initialize session tracking
    if (-not $Global:__MtSession) {
        $Global:__MtSession = @{}
    }
    $Global:__MtSession.Connections = $Service

    # Connect to Microsoft Graph
    if ($Service -contains 'Graph' -or $Service -contains 'All') {
        Write-Verbose 'Connecting to Microsoft Graph'

        try {
            $graphScopes = Get-MtGraphScope -SendMail:$SendMail -SendTeamsMessage:$SendTeamsMessage -Privileged:$Privileged

            switch ($PSCmdlet.ParameterSetName) {
                'Interactive' {
                    if ($TenantId) {
                        Connect-MgGraph -Scopes $graphScopes -NoWelcome -Environment $Environment -TenantId $TenantId
                    } else {
                        Connect-MgGraph -Scopes $graphScopes -NoWelcome -Environment $Environment
                        $TenantId = (Get-MgContext).TenantId
                    }
                }
                'DeviceCode' {
                    if ($TenantId) {
                        Connect-MgGraph -Scopes $graphScopes -NoWelcome -UseDeviceCode -Environment $Environment -TenantId $TenantId
                    } else {
                        Connect-MgGraph -Scopes $graphScopes -NoWelcome -UseDeviceCode -Environment $Environment
                        $TenantId = (Get-MgContext).TenantId
                    }
                }
                'ManagedIdentity' {
                    Connect-MgGraph -Identity -NoWelcome
                    $context = Get-MgContext
                    $TenantId = $context.TenantId
                    Write-Host "Connected to Microsoft Graph using Managed Identity" -ForegroundColor Green
                }
                'ClientSecret' {
                    $clientSecretCredential = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
                    Connect-MgGraph -ClientSecretCredential $clientSecretCredential -TenantId $TenantId -NoWelcome
                    Write-Host "Connected to Microsoft Graph using Client Secret" -ForegroundColor Green
                }
                'Certificate' {
                    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome
                    Write-Host "Connected to Microsoft Graph using Certificate (Thumbprint)" -ForegroundColor Green
                }
                'CertificateFile' {
                    if ($CertificatePassword) {
                        $cert = Get-PfxCertificate -FilePath $CertificatePath -Password $CertificatePassword
                    } else {
                        $cert = Get-PfxCertificate -FilePath $CertificatePath
                    }
                    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $cert -NoWelcome
                    Write-Host "Connected to Microsoft Graph using Certificate File" -ForegroundColor Green
                }
            }

            $context = Get-MgContext
            if ($context) {
                Write-Host "Successfully connected to Microsoft Graph!" -ForegroundColor Green
                Write-Host "  Tenant: $($context.TenantId)" -ForegroundColor Gray
                Write-Host "  Account: $($context.Account)" -ForegroundColor Gray
                Write-Host "  Environment: $($context.Environment)" -ForegroundColor Gray
                Write-Host "  Scopes: $($context.Scopes -join ', ')" -ForegroundColor Gray
            }

        } catch [System.Management.Automation.CommandNotFoundException] {
            Write-Error "The Microsoft Graph PowerShell module is not installed. Please install the module using: Install-Module Microsoft.Graph -Scope CurrentUser"
            return
        } catch {
            Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
            return
        }
    }

    # Connect to Azure PowerShell
    if ($Service -contains 'Azure' -or $Service -contains 'All') {
        Write-Verbose 'Connecting to Azure'

        try {
            switch ($PSCmdlet.ParameterSetName) {
                'Interactive' {
                    if ($TenantId) {
                        Connect-AzAccount -SkipContextPopulation -Environment $AzureEnvironment -TenantId $TenantId
                    } else {
                        Connect-AzAccount -SkipContextPopulation -Environment $AzureEnvironment
                    }
                }
                'DeviceCode' {
                    if ($TenantId) {
                        Connect-AzAccount -SkipContextPopulation -UseDeviceAuthentication -Environment $AzureEnvironment -TenantId $TenantId
                    } else {
                        Connect-AzAccount -SkipContextPopulation -UseDeviceAuthentication -Environment $AzureEnvironment
                    }
                }
                'ManagedIdentity' {
                    Connect-AzAccount -Identity -SkipContextPopulation
                    Write-Host "Connected to Azure using Managed Identity" -ForegroundColor Green
                }
                'ClientSecret' {
                    $azCredential = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
                    Connect-AzAccount -ServicePrincipal -Credential $azCredential -TenantId $TenantId -Environment $AzureEnvironment
                    Write-Host "Connected to Azure using Service Principal (Client Secret)" -ForegroundColor Green
                }
                'Certificate' {
                    Connect-AzAccount -ServicePrincipal -ApplicationId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -Environment $AzureEnvironment
                    Write-Host "Connected to Azure using Service Principal (Certificate)" -ForegroundColor Green
                }
                'CertificateFile' {
                    Connect-AzAccount -ServicePrincipal -ApplicationId $ClientId -TenantId $TenantId -CertificatePath $CertificatePath -Environment $AzureEnvironment
                    Write-Host "Connected to Azure using Service Principal (Certificate File)" -ForegroundColor Green
                }
            }

        } catch [System.Management.Automation.CommandNotFoundException] {
            Write-Warning "The Azure PowerShell module (Az) is not installed. Please install using: Install-Module Az -Scope CurrentUser"
        } catch {
            Write-Error "Failed to connect to Azure: $($_.Exception.Message)"
        }
    }

    # Connect to Exchange Online
    if ($Service -contains 'ExchangeOnline' -or $Service -contains 'All') {
        Write-Verbose 'Connecting to Exchange Online'

        try {
            switch ($PSCmdlet.ParameterSetName) {
                'Interactive' {
                    if ($TenantId) {
                        Connect-ExchangeOnline -ExchangeEnvironmentName $ExchangeEnvironmentName -ShowBanner:$false -Organization $TenantId
                    } else {
                        Connect-ExchangeOnline -ExchangeEnvironmentName $ExchangeEnvironmentName -ShowBanner:$false
                    }
                }
                'ManagedIdentity' {
                    Connect-ExchangeOnline -ManagedIdentity -Organization $TenantId -ShowBanner:$false
                    Write-Host "Connected to Exchange Online using Managed Identity" -ForegroundColor Green
                }
                'ClientSecret' {
                    # Note: Exchange Online PowerShell doesn't directly support client secret authentication
                    # This would require a custom implementation or using certificate-based authentication
                    Write-Warning "Exchange Online PowerShell doesn't support client secret authentication directly. Consider using certificate-based authentication instead."
                }
                'Certificate' {
                    Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $TenantId -ShowBanner:$false
                    Write-Host "Connected to Exchange Online using Certificate" -ForegroundColor Green
                }
                'CertificateFile' {
                    Connect-ExchangeOnline -AppId $ClientId -CertificateFilePath $CertificatePath -Organization $TenantId -ShowBanner:$false
                    Write-Host "Connected to Exchange Online using Certificate File" -ForegroundColor Green
                }
                default {
                    Connect-ExchangeOnline -ExchangeEnvironmentName $ExchangeEnvironmentName -ShowBanner:$false
                }
            }

        } catch [System.Management.Automation.CommandNotFoundException] {
            Write-Warning "The Exchange Online PowerShell module is not installed. Please install using: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
        } catch {
            Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
        }
    }

    # Output connection summary
    Write-Host "`nConnection Summary:" -ForegroundColor Cyan
    Write-Host "Services: $($Service -join ', ')" -ForegroundColor White
    Write-Host "Authentication Method: $($PSCmdlet.ParameterSetName)" -ForegroundColor White
    if ($TenantId) {
        Write-Host "Tenant ID: $TenantId" -ForegroundColor White
    }
    if ($ClientId) {
        Write-Host "Client ID: $ClientId" -ForegroundColor White
    }
    Write-Host "Environment: $Environment" -ForegroundColor White
}

# Helper function to create a new registered application for Maester
function New-MaesterAppRegistration {
    <#
    .SYNOPSIS
    Creates a new Azure AD app registration for Maester with required permissions.

    .DESCRIPTION
    This function creates a new Azure AD application registration with the necessary API permissions
    for Maester to function properly. It can also generate a client secret or certificate for authentication.

    .EXAMPLE
    New-MaesterAppRegistration -DisplayName "Maester Security Testing" -UseClientSecret

    Creates a new app registration with a client secret.

    .EXAMPLE
    New-MaesterAppRegistration -DisplayName "Maester Security Testing" -UseCertificate -CertificateSubject "CN=MaesterApp"

    Creates a new app registration with a self-signed certificate.
    #>
    [CmdletBinding()]
    param(
        # The display name for the application registration
        [Parameter(Mandatory = $true)]
        [string] $DisplayName,

        # Create a client secret for the application
        [switch] $UseClientSecret,

        # Use certificate-based authentication
        [switch] $UseCertificate,

        # The certificate subject name (only used when UseCertificate is specified)
        [string] $CertificateSubject = "CN=MaesterApp",

        # The validity period for the client secret or certificate in months
        [int] $ValidityMonths = 12
    )

    # Check if connected to Microsoft Graph
    $context = Get-MgContext
    if (-not $context) {
        Write-Error "Please connect to Microsoft Graph first using Connect-MgGraph or Connect-Maester"
        return
    }

    try {
        # Get required Graph scopes for Maester
        $requiredScopes = Get-MtGraphScope -Privileged

        Write-Host "Creating Azure AD application registration..." -ForegroundColor Cyan

        # Create the application registration
        $app = New-MgApplication -DisplayName $DisplayName -SignInAudience "AzureADMyOrg"

        Write-Host "Application created successfully!" -ForegroundColor Green
        Write-Host "Application ID: $($app.AppId)" -ForegroundColor White
        Write-Host "Object ID: $($app.Id)" -ForegroundColor White

        # Add required API permissions
        Write-Host "Adding required API permissions..." -ForegroundColor Cyan

        $graphServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"

        $requiredResourceAccess = @{
            ResourceAppId = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph
            ResourceAccess = @()
        }

        foreach ($scope in $requiredScopes) {
            $permission = $graphServicePrincipal.AppRoles | Where-Object { $_.Value -eq $scope }
            if ($permission) {
                $requiredResourceAccess.ResourceAccess += @{
                    Id = $permission.Id
                    Type = "Role"
                }
            }
        }

        Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @($requiredResourceAccess)

        # Create service principal
        $servicePrincipal = New-MgServicePrincipal -AppId $app.AppId

        Write-Host "Service Principal created: $($servicePrincipal.Id)" -ForegroundColor Green

        # Handle authentication method
        if ($UseClientSecret) {
            Write-Host "Creating client secret..." -ForegroundColor Cyan

            $endDate = (Get-Date).AddMonths($ValidityMonths)
            $passwordCredential = @{
                DisplayName = "Maester Client Secret"
                EndDateTime = $endDate
            }

            $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $passwordCredential

            Write-Host "Client secret created successfully!" -ForegroundColor Green
            Write-Host "Client Secret: $($secret.SecretText)" -ForegroundColor Yellow
            Write-Host "Secret ID: $($secret.KeyId)" -ForegroundColor White
            Write-Host "Expires: $($secret.EndDateTime)" -ForegroundColor White
            Write-Warning "Please save the client secret value securely. It will not be shown again!"
        }

        if ($UseCertificate) {
            Write-Host "Creating self-signed certificate..." -ForegroundColor Cyan

            $cert = New-SelfSignedCertificate -Subject $CertificateSubject -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddMonths($ValidityMonths)

            # Export certificate
            $certPath = Join-Path $env:TEMP "$($app.AppId)_certificate.cer"
            Export-Certificate -Cert $cert -FilePath $certPath | Out-Null

            # Read certificate and convert to base64
            $certBytes = Get-Content $certPath -AsByteStream
            $certBase64 = [System.Convert]::ToBase64String($certBytes)

            # Add certificate to application
            $keyCredential = @{
                Type = "AsymmetricX509Cert"
                Usage = "Verify"
                DisplayName = "Maester Certificate"
                Key = [System.Text.Encoding]::UTF8.GetBytes($certBase64)
                EndDateTime = $cert.NotAfter
            }

            Update-MgApplication -ApplicationId $app.Id -KeyCredentials @($keyCredential)

            Write-Host "Certificate created and added to application!" -ForegroundColor Green
            Write-Host "Certificate Thumbprint: $($cert.Thumbprint)" -ForegroundColor White
            Write-Host "Certificate Path: $certPath" -ForegroundColor White
            Write-Host "Certificate Store Location: Cert:\CurrentUser\My\$($cert.Thumbprint)" -ForegroundColor White
            Write-Host "Expires: $($cert.NotAfter)" -ForegroundColor White
        }

        # Output connection instructions
        Write-Host "`nTo connect Maester using this application:" -ForegroundColor Cyan

        if ($UseClientSecret) {
            Write-Host "Connect-Maester -ClientId '$($app.AppId)' -TenantId '$($context.TenantId)' -ClientSecret (ConvertTo-SecureString 'YOUR_CLIENT_SECRET' -AsPlainText -Force)" -ForegroundColor Gray
        }

        if ($UseCertificate) {
            Write-Host "Connect-Maester -ClientId '$($app.AppId)' -TenantId '$($context.TenantId)' -CertificateThumbprint '$($cert.Thumbprint)'" -ForegroundColor Gray
        }

        Write-Warning "Don't forget to grant admin consent for the API permissions in the Azure portal!"

        return @{
            ApplicationId = $app.AppId
            ObjectId = $app.Id
            TenantId = $context.TenantId
            ServicePrincipalId = $servicePrincipal.Id
            DisplayName = $DisplayName
        }

    } catch {
        Write-Error "Failed to create application registration: $($_.Exception.Message)"
        return $null
    }
}

# Helper function to grant admin consent for Maester app
function Grant-MaesterAppConsent {
    <#
    .SYNOPSIS
    Grants admin consent for a Maester application registration.

    .DESCRIPTION
    This function grants admin consent for the API permissions of a Maester application registration.

    .EXAMPLE
    Grant-MaesterAppConsent -ApplicationId "12345678-1234-1234-1234-123456789012"

    Grants admin consent for the specified application.
    #>
    [CmdletBinding()]
    param(
        # The Application ID of the registered application
        [Parameter(Mandatory = $true)]
        [string] $ApplicationId
    )

    try {
        # Get the service principal for the application
        $servicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$ApplicationId'"

        if (-not $servicePrincipal) {
            Write-Error "Service principal not found for application ID: $ApplicationId"
            return
        }

        # Get Microsoft Graph service principal
        $graphServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"

        # Get the required permissions
        $requiredScopes = Get-MtGraphScope -Privileged

        Write-Host "Granting admin consent for API permissions..." -ForegroundColor Cyan

        foreach ($scope in $requiredScopes) {
            $permission = $graphServicePrincipal.AppRoles | Where-Object { $_.Value -eq $scope }

            if ($permission) {
                try {
                    # Check if consent already exists
                    $existingGrant = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id | Where-Object { $_.AppRoleId -eq $permission.Id }

                    if (-not $existingGrant) {
                        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id -PrincipalId $servicePrincipal.Id -AppRoleId $permission.Id -ResourceId $graphServicePrincipal.Id
                        Write-Host "  ✓ Granted: $scope" -ForegroundColor Green
                    } else {
                        Write-Host "  ✓ Already granted: $scope" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Warning "  ✗ Failed to grant: $scope - $($_.Exception.Message)"
                }
            }
        }

        Write-Host "Admin consent process completed!" -ForegroundColor Green

    } catch {
        Write-Error "Failed to grant admin consent: $($_.Exception.Message)"
    }
}