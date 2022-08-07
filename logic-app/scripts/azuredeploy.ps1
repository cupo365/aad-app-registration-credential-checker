Param(
    [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $TenantId,
    [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $SubscriptionId,
    [parameter(Mandatory = $false, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName = "rg-aadappregcredentialcheck",
    [parameter(Mandatory = $false, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupLocation = "West Europe",
    [parameter(Mandatory = $false, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $AzureResourceTemplateFilePath = "../arm/azuredeploy.json",
    [parameter(Mandatory = $false, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $AzureResourceTemplateParametersFilePath = "../arm/azuredeploy.parameters.production.json",
    [parameter(Mandatory = $false, ValueFromPipeline = $true)] [array] $Modules = ("AzureAD", "Az", "AzureRm"), #"AzureRm"
    [parameter(Mandatory = $false, ValueFromPipeline = $true)] [array] $MsiPermissions = @( 
        [pscustomobject]@{ServicePrincipalId = "00000003-0000-0000-c000-000000000000"; Scope = "User.Read.All" }  
        [pscustomobject]@{ServicePrincipalId = "00000003-0000-0000-c000-000000000000"; Scope = "Application.Read.All" }
    ),
    [parameter(Mandatory = $false, ValueFromPipeline = $true)] [string] $UserName,
    [parameter(Mandatory = $false, ValueFromPipeline = $true)] [string] $Pass
)

<#
    .SYNOPSIS
        Script to deploy the Azure resources for the AAD App Reg Credential Check solution.

    .DESCRIPTION
        This script allows you to deploy the AAD App Reg Credential Check solution almost fully automated.
        The script will create a resource group, deploy and configure the solution.

    .PARAMETER TenantId <string> [required]
        The id of the tenant.
    
    .PARAMETER SubscriptionId <string> [required]
        The id of the Azure subscription to which the Azure resources should be assigned to.
    
    .PARAMETER ResourceGroupName <string> [optional]
        The name of the Azure resource group in which the Azure resources should be deployed.
        Default value is 'rg-aadappregcredentialcheck'.
    
    .PARAMETER ResourceGroupLocation <string> [optional]
        The location of the Azure resource group.
        Default value is 'West Europe'.
    
    .PARAMETER AzureResourceTemplateFilePath <string> [optional]
        The (relative) file path to the Azure Resource Manager template to deploy.
        Default value is '../arm/azuredeploy.json'.
    
    .PARAMETER AzureResourceTemplateParametersFilePath <string> [optional]
        The (relative) file path to the Azure Resource Manager template parameters to deploy.
        Default value is '../arm/azuredeploy.parameters.production.json'.

    .PARAMETER Modules <array> [optional]
        The required PowerShell modules to execute the program.
        Default value is "AzureAD", "Az", "AzureRm" (minimum required).

    .PARAMETER MsiPermissionScopes <array> [optional]
        Specify one or more permissions to add to the managed identity.
        For example: "Directory.Read.All", "Group.Read.All", "GroupMember.Read.All", "User.Read.All" or "Sites.FullControl.All".
        Default value is "User.Read.All", "Application.Read.All" (minimum required).

    .PARAMETER ServicePrincipalId <string> [optional]
        The app id of the service principal from which you want to add permissions to the managed identity.
        Caution: object id and app ip are different fields/values!
        Some examples:
        - The Graph API has the following id: 00000003-0000-0000-c000-000000000000
        - The SharePoint service principal has the following id: 00000003-0000-0ff1-ce00-000000000000
        Default value is '00000003-0000-0000-c000-000000000000' (minimum required).
    
    .PARAMETER UserName <string> [optional]
        The username of the account to login to Azure with.
    
    .PARAMETER Pass <string> [optional]
        The password of the account to login to Azure with.

    .LINK
        Inspired by: 
        - https://aztoso.com/security/microsoft-graph-permissions-managed-identity/
        - https://github.com/logicappsio/LogicAppConnectionAuth/blob/master/LogicAppConnectionAuth.ps1
#>

<# ---------------- Global variables ---------------- #>
$GLOBAL:requiredPSVersion = "5.1"
$GLOBAL:creds = ""

<# ---------------- Helper functions ---------------- #>
Function Initialize {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [array] $Modules,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $TenantId,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $SubscriptionId,
        [parameter(Mandatory = $false, ValueFromPipeline = $true)] [AllowEmptyString()] [string] $UserName,
        [parameter(Mandatory = $false, ValueFromPipeline = $true)] [AllowEmptyString()] [string] $Pass
    )

    Try {
        Validate-PSVersion $GLOBAL:requiredPSVersion

        Write-Host "`n|-----------------------[ (1/6) Importing modules ]-----------------------|" -ForegroundColor Cyan
        Foreach ($Module in $Modules) {
            Import-PSModule $Module
        }

        Write-Host "`n|----------------------[ (2/6) Connecting to Azure ]----------------------|" -ForegroundColor Cyan
        Write-Host "`nConnecting to Azure..." -ForegroundColor Magenta
        If ($false -eq [string]::IsNullOrEmpty($UserName) -And $false -eq [string]::IsNullOrEmpty($Pass)) {
            $GLOBAL:creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $(convertto-securestring $Pass -AsPlainText -Force)
        }
        Else {
            $GLOBAL:creds = Get-Credential
        }
        Connect-AzAccount -Tenant $TenantId -SubscriptionId $SubscriptionId -Credential $GLOBAL:creds | Out-Null
        Write-Host "Connected!" -ForegroundColor Green
    }
    Catch [Exception] {
        Write-Host "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while initializing. Message: $($_.Exception.Message)" -ForegroundColor Yellow

        Write-Host "Terminating program. Reason: could not initialize." -ForegroundColor Red
        Finish-Up
    }
}

Function Validate-PSVersion {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $RequiredPSVersion
    )

    Try {
        # Validate PowerShell version
        # There is an issue with using PowerShell 7 with the AzureAD module, therefore force v5.1
        # More info here: https://github.com/PowerShell/PowerShell/issues/10473
        if ($false -eq $PSVersionTable.PSVersion.ToString().StartsWith($GLOBAL:requiredPSVersion)) {
            # Re-launch as configured PS version if the script is not running the required version
            powershell -Version $RequiredPSVersion -File $PSCommandPath
            exit
        }
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred on line $($_.InvocationInfo.ScriptLineNumber) while validating the PowerShell version. Message: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Terminating program. Reason: could not validate PowerShell version." -ForegroundColor Red
        Finish-Up
    }
}

Function Import-PSModule {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ModuleName,
        [parameter(Mandatory = $false, ValueFromPipeline = $true)] [boolean] $AttemptedBefore
    )

    Try {
        # Import module
        Write-Host "`nImporting $($ModuleName) module..." -ForegroundColor Magenta
        Import-Module -Name $ModuleName -Scope Local -Force -ErrorAction Stop | Out-Null        
        Write-Host "Successfully imported $($ModuleName)!" -ForegroundColor Green
    }
    Catch [Exception] {
        Write-Host "$($ModuleName) was not found on the specified location." -ForegroundColor Yellow

        If ($true -eq $AttemptedBefore) {
            Write-Host "Terminating program. Reason: could not import dependend modules." -ForegroundColor Red
            Finish-Up
        }
        Else {            
            Install-PSModule $ModuleName
        }
    }
}

Function Install-PSModule {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ModuleName
    )

    Try {
        # Install module
        Write-Host "`nInstalling $($ModuleName) module..." -ForegroundColor Magenta
        Install-Module -Name $ModuleName -Scope CurrentUser -AllowClobber -Force -ErrorAction Stop | Out-Null
        Write-Host "Successfully installed $($ModuleName)!" -ForegroundColor Green

        Import-PSModule $ModuleName $true
    }
    Catch [Exception] {
        Write-Host "Could not install $($ModuleName)." -ForegroundColor Yellow
        
        Write-Host "Terminating program. Reason: could not import dependend modules." -ForegroundColor Red
        Finish-Up
    }
}

Function Verify-ResourceGroup {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupLocation
    )

    Try {
        Write-Host "`n|--------------------[ (3/6) Verify Resource Group ]--------------------|" -ForegroundColor Cyan        
        # Validate resource group existence
        Write-Host "`nVerifying the existence of resource group '$($ResourceGroupName)'..." -ForegroundColor Magenta
        Get-AzResourceGroup -Name $ResourceGroupName -ErrorVariable notPresent -ErrorAction SilentlyContinue | Out-Null

        if ($notPresent) {
            # Resource group doesn't exist yet
            Write-Host "Resource group does not exist. Creating..." -ForegroundColor Yellow
            $NewResourceGroup = New-AzResourceGroup -Name $ResourceGroupName -Location $ResourceGroupLocation

            If ("Succeeded" -eq $NewResourceGroup.ProvisioningState) {
                Write-Host "Created!" -ForegroundColor Green
            }
            Else {
                throw "Failed to create resource group with name ($ResourceGroupName) and location ($ResourceGroupLocation). Provisioning state: $($NewResourceGroup.ProvisioningState)."
            }
        }
        else {
            # Resource group already exists
            Write-Host "Resource group already exists on the tenant." -ForegroundColor Green
        }
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred on line $($_.InvocationInfo.ScriptLineNumber) while validating the resource group. Message: $($_.Exception.Message)" -ForegroundColor Yellow

        Write-Host "Terminating program. Reason: encountered an error before program could successfully finish." -ForegroundColor Red
        Finish-Up
    }
}

Function Deploy-AzureResourceTemplate {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $AzureResourceTemplateFilePath,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $AzureResourceTemplateParametersFilePath
    )

    Try {
        Write-Host "`n|----------------------[ (4/6) Deploy ARM Template ]----------------------|" -ForegroundColor Cyan
        Write-Host "`nDeploying Azure Resource Template..." -ForegroundColor Magenta
        $Deployment = New-AzResourceGroupDeployment -Name "PowerShellDeploy-$(Get-Date -Format "HHmmddMMyyyy")" -ResourceGroupName $ResourceGroupName -TemplateFile $AzureResourceTemplateFilePath -TemplateParameterFile $AzureResourceTemplateParametersFilePath

        Write-Host "Verifying the deployment..." -ForegroundColor Yellow
        If ("Succeeded" -eq $Deployment.ProvisioningState -And $null -ne $Deployment.Outputs) {
            Write-Host "Deployed!" -ForegroundColor Green
        }
        Else {
            throw "Failed to deploy the Azure Resource Template. Provisioning state: $($Deployment.ProvisioningState)."
        }

        return $Deployment.Outputs
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred on line $($_.InvocationInfo.ScriptLineNumber) while deploying the ARM template. Message: $($_.Exception.Message)" -ForegroundColor Yellow

        Write-Host "Terminating program. Reason: encountered an error before program could successfully finish." -ForegroundColor Red
        Finish-Up
    }
}

Function Add-MSIPermissions {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $TenantId,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $MsiObjectId,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [array] $MsiPermissions
    )

    Try {
        Write-Host "`n|-------------------[ (5/6) Configure MSI Permissions ]-------------------|" -ForegroundColor Cyan  
        # Connect to Azure AD
        Write-Host "`nConnecting to Azure Active Directory..." -ForegroundColor Magenta
        Connect-AzureAD -TenantId $TenantId -Credential $GLOBAL:creds | Out-Null
        Write-Host "Connected!" -ForegroundColor Green

        # Configure MSI permissions
        Write-Host "`nChecking if MSI with ID '$($MsiObjectId)' exists..." -ForegroundColor Magenta
        $LoopCount = 0
        $Msi = $null

        Do {
            $LoopCount++

            Try {
                $Msi = Get-AzureADServicePrincipal -Filter "objectId eq '$($MsiObjectId)'"
            }
            Catch [Exception] {
                # Give deployment some slack
                Write-Host "`nMSI does not exist yet. Waiting 10 seconds to give deployment some slack..." -ForegroundColor Magenta
                Start-Sleep -Seconds 10
                Write-Host "Resuming. Checking again (attempt $($LoopCount + 1) out of maximum 6)..." -ForegroundColor Yellow                
            }
            
        } Until ($null -ne $Msi -Or 5 -lt $LoopCount)

        If ($null -eq $Msi) {
            throw "Cannot fetch MSI with ID '$($MsiObjectId)' on tenant."
        }
        Write-Host "MSI exists!" -ForegroundColor Green

        $CurrentMsiPermissionScopes = Get-AzureADServiceAppRoleAssignedTo -ObjectId $MsiObjectId -All $true

        Foreach ($MsiPermission in $MsiPermissions) {
            Write-Host "`nChecking if service principal with ID '$($MsiPermission.ServicePrincipalId)' exists..." -ForegroundColor Magenta
            $ServicePrincipal = Get-AzureADServicePrincipal -Filter "appId eq '$($MsiPermission.ServicePrincipalId)'" -ErrorAction SilentlyContinue

            If ($null -eq $ServicePrincipal) {
                Write-Host "Cannot fetch service principal for permission scope '$($MsiPermission.Scope)'. This permission will be skipped. Continuing." -ForegroundColor Yellow
                continue
            }
            Write-Host "Service principal exists!" -ForegroundColor Green

            Write-Host "Fetching permission scope '$($MsiPermission.Scope)'..." -ForegroundColor Magenta
            $PermissionScope = $ServicePrincipal.AppRoles | Where-Object { $_.Value -eq $MsiPermission.Scope -and $_.AllowedMemberTypes -contains "Application" }

            If ($null -eq $PermissionScope) {
                throw "Could not fetch permission scope '$($MsiPermission.Scope)' for service principal with ID '$($MsiPermission.ServicePrincipalId)'."
            }

            $ExistingMsiPermissionScope = $CurrentMsiPermissionScopes | Where-Object { $_.Id -eq $PermissionScope.Id }
            If ($null -ne $ExistingMsiPermissionScope) {
                Write-Host "Permission already exists on the MSI. Continuing." -ForegroundColor Yellow
                continue
            }
            Write-Host "Fetched!" -ForegroundColor Green

            Write-Host "Adding permission scope '$($MsiPermission.Scope)' of service principal '$($MsiPermission.ServicePrincipalId)' to MSI '$($MsiObjectId)'..." -ForegroundColor Magenta
            New-AzureAdServiceAppRoleAssignment -ObjectId $Msi.ObjectId -PrincipalId $Msi.ObjectId -ResourceId $ServicePrincipal.ObjectId -Id $PermissionScope.Id -ErrorAction Stop | Out-Null
            Write-Host "Added!" -ForegroundColor Green
        }
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred on line $($_.InvocationInfo.ScriptLineNumber) while configuring the MSI permission scopes. Message: $($_.Exception.Message)" -ForegroundColor Yellow

        Write-Host "Terminating program. Reason: encountered an error before program could successfully finish." -ForegroundColor Red
        Finish-Up
    }
}

Function Verify-Connection {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $TenantId,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ConnectionName
    )

    Try {
        Write-Host "`n|-----------------------[ (6/6) Verify Connection ]-----------------------|" -ForegroundColor Cyan
        # Prevent get_SerializationSettings error by enabling AzureRm prefix aliases for Az modules.
        # More information here: https://azurelessons.com/method-get_serializationsettings-error/
        Enable-AzureRmAlias -Scope CurrentUser

        # Connect to Azure AD
        Write-Host "`nConnecting to Azure Remote Account..." -ForegroundColor Magenta
        Login-AzureRmAccount -TenantId $TenantId -Credential $GLOBAL:creds | Out-Null
        Write-Host "Connected!" -ForegroundColor Green

        # Verify connection        
        $Connection = Get-ConnectionStatus $ResourceGroupName $ConnectionName

        If ($null -eq $Connection) {
            throw "Could not fetch connection '$($ConnectionName)'."
        }

        If ("Connected" -ne $Connection.Properties.Statuses[0].status) {
            Write-Host "`nConnection is not authorized yet. Current status: $($Connection.Properties.Statuses[0].status). Attempting connection authorization..." -ForegroundColor Yellow

            Authorize-Connection $Connection.ResourceId

            $Connection = Get-ConnectionStatus $ResourceGroupName $ConnectionName

            If ($null -eq $Connection) {
                throw "Could not fetch connection '$($ConnectionName)'."
            }

            If ("Connected" -ne $Connection.Properties.Statuses[0].status) {
                throw "Could not verify authorized connection '$($ConnectionName)'."
            }

            Write-Host "Connection '$($ConnectionName)' is authorized!" -ForegroundColor Green
        }
        Else {
            Write-Host "Connection '$($ConnectionName)' is already authorized." -ForegroundColor Yellow
        }
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred on line $($_.InvocationInfo.ScriptLineNumber) while verifying the connection. Message: $($_.Exception.Message)" -ForegroundColor Yellow

        Write-Host "Terminating program. Reason: encountered an error before program could successfully finish." -ForegroundColor Red
        Finish-Up
    }
}
Function Show-OAuthWindow {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $AuthorizationUrl
    )

    Try {
        Write-Host "`nAssembling authorization dialog..." -ForegroundColor Magenta
        Add-Type -AssemblyName System.Windows.Forms
 
        $Form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width = 600; Height = 800 }
        $Web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width = 580; Height = 780; Url = ($AuthorizationUrl -f ($Scope -join "%20")) }
        $DocComp = {
            $Global:uri = $Web.Url.AbsoluteUri
            if ($Global:Uri -match "error=[^&]*|code=[^&]*") { 
                $Form.Close() 
            }
        }
        $Web.ScriptErrorsSuppressed = $true
        $Web.Add_DocumentCompleted($DocComp)
        $Form.Controls.Add($Web)
        $Form.Add_Shown({ $Form.Activate() })

        Write-Host "Assembled! Pending user authorization input..." -ForegroundColor Yellow
        $Form.ShowDialog() | Out-Null
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred on line $($_.InvocationInfo.ScriptLineNumber) while assembling the authorization dialog. Message: $($_.Exception.Message)" -ForegroundColor Yellow

        Write-Host "Terminating program. Reason: encountered an error before program could successfully finish." -ForegroundColor Red
        Finish-Up
    }
}

Function Get-ConnectionStatus {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName,
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ConnectionName
    )

    Try {
        Write-Host "`nFetching connection '$($ConnectionName)'..." -ForegroundColor Magenta
        $Connection = Get-AzureRmResource -ResourceType "Microsoft.Web/connections" -ResourceGroupName $ResourceGroupName -ResourceName $ConnectionName -ErrorAction SilentlyContinue

        If ($null -eq $Connection) {
            throw "Could not fetch connection."
        }
        
        Write-Host "Fetched!" -ForegroundColor Green

        return $Connection
    }
    Catch [Exception] {
        return $null
    }
}

Function Authorize-Connection {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)] [ValidateNotNullOrEmpty()] [string] $ConnectionResourceId
    )

    Try {
        Write-Host "Fetching connection authorization link..." -ForegroundColor Magenta

        $Parameters = @{
            "parameters" = , @{
                "parameterName" = "token";
                "redirectUrl"   = "https://ema1.exp.azure.com/ema/default/authredirect"
            }
        }

        $Response = Invoke-AzureRmResourceAction -Action "listConsentLinks" -ResourceId $ConnectionResourceId -Parameters $Parameters -Force -ErrorAction SilentlyContinue
        
        If ($null -eq $Response -Or $null -eq $Response.Value.Link) {
            throw "Could not fetch connection authorization link."
        }
        Write-Host "Fetched!" -ForegroundColor Green
        
        $AuthorizationUrl = $Response.Value.Link
        Show-OAuthWindow $AuthorizationUrl

        $Code = ($uri | Select-string -pattern '(code=)(.*)$').Matches[0].Groups[2].Value
        If ([string]::IsNullOrEmpty($Code)) {
            throw "Could not fetch access code from user authorization."
        }
        Write-Host "Successfully fetched an access code!" -ForegroundColor Green

        $Parameters = @{ }
        $Parameters.Add("code", $Code)
        
        Write-Host "`nAuthorizing connection with access code..." -ForegroundColor Magenta        
        Invoke-AzureRmResourceAction -Action "confirmConsentCode" -ResourceId $ConnectionResourceId -Parameters $Parameters -Force -ErrorAction SilentlyContinue -ErrorVariable error | Out-Null

        # There is a known bug in PowerShell which throws a meaningless error about an argument 'obj' being null and not being able to process it.
        # This error does not mean the action went wrong. Therefore, it should not be declared as a breaking error
        # More information here: https://social.technet.microsoft.com/Forums/windowsserver/en-US/f0edd01c-7711-46cd-b47b-45785d44f062/why-is-the-value-of-argument-quotobjquot-null
        If ('Cannot process argument because the value of argument "obj" is null. Change the value of argument "obj" to a non-null value.' -ne $error) {
            throw "Could not authorize connection."
        }
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred on line $($_.InvocationInfo.ScriptLineNumber) while assembling the authorization dialog. Message: $($_.Exception.Message)" -ForegroundColor Yellow

        Write-Host "Terminating program. Reason: encountered an error before program could successfully finish." -ForegroundColor Red
        Finish-Up
    }
}

Function Finish-Up {
    Try {
        Write-Host "`n|----------------------------[ Finishing up ]-----------------------------|" -ForegroundColor Cyan

        Write-Host "`nDisconnecting the session..." -ForegroundColor Magenta
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        Disconnect-AzureAD -ErrorAction SilentlyContinue | Out-Null
        Logout-AzureRmAccount -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Disconnected!" -ForegroundColor Green

        Write-Host "`n|------------------------------[ Finished ]-------------------------------|`n" -ForegroundColor Cyan

        exit 1
    }
    Catch [Exception] {
        Write-Host "`n|------------------------------[ Finished ]-------------------------------|`n" -ForegroundColor Cyan
        exit 1
    }
}

<# ---------------- Program execution ---------------- #>
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
# $env:PSModulePath = $env:PSModulePath + ';{PATH_TO_MODULE_DIRECTORY}'

Try {
    Initialize $Modules $TenantId $SubscriptionId $UserName $Pass

    Verify-ResourceGroup $ResourceGroupName $ResourceGroupLocation

    $DeploymentOutputs = Deploy-AzureResourceTemplate $ResourceGroupName $AzureResourceTemplateFilePath $AzureResourceTemplateParametersFilePath

    Add-MSIPermissions $TenantId $DeploymentOutputs.logicAppSystemAssignedIdentityObjectId.Value $MsiPermissions

    $DeployedConnections = @($DeploymentOutputs.office365ConnectionName.Value)

    Foreach ($DeployedConnection in $DeployedConnections) {
        Verify-Connection $TenantId $ResourceGroupName $DeployedConnection
    }

    Write-Host "`nSuccessfully executed the program!" -ForegroundColor Green

    Finish-Up
}
Catch [Exception] {
    Write-Host "`nAn error occurred on line $($_.InvocationInfo.ScriptLineNumber) while initializing. Message: $($_.Exception.Message)" -ForegroundColor Yellow

    Write-Host "Terminating program. Reason: encountered an error before program could successfully finish." -ForegroundColor Red

    Finish-Up
}
