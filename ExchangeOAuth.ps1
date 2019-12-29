<#  Script:     PowerShell Exchange OAuth
#   Author:     Kai Burton
#   TechNet:    https://social.technet.microsoft.com/profile/-kai/
#   Sources:    {ExoPowershellModule.dll, Microsoft.ADAL.PowerShell.psm1}
#   Comments:   Admin rights should not be required to run this script, AccessTokens can also be stored & re-used until expiration 
                RefreshToken can be used to generate a new AccessToken
#>

<# Script varibles #>
$script:AuthorizationToken = $null
$script:UserPrincipalName = $null
$script:ReadyScript = $false
$script:args = $MyInvocation.Line

<# Preset varibles #>
$script:AppId = "a0c73c16-a7e3-4564-9a95-2bdf47383716"
$script:ResourceId = "https://outlook.office365.com"
$script:EndpointURI = "https://login.windows.net/common"
$script:ConnectionURI = "https://outlook.office365.com:443/PowerShell-LiveId?BasicAuthToOAuthConversion=true"
$script:RedirectURI = "urn:ietf:wg:oauth:2.0:oob"
$script:AuthorityName = "portal.ms"

<#  Type:   Function-Void
#   Desc:   Calls Functions: {Initialize-AccessToken, Assert-ReadyADALModule} | Connects to Exchange Online, using AccessTokens & Imports PSModule
#   Usage:  Connect-ExchangeOnlineOAuth
#   Ref:    NONE - Main Void #>
Function Connect-ExchangeOnlineOAuth {

    # Single thread while-loop - "should not" cause any performance or script issues
    # Attempt to install the ADAL Module 3 times, if failed - user/admin needs to install the Adal Module manually
    $InstallationAttempts = 0
    while(!($script:ReadyScript)) {
        (Assert-ReadyADALModule | Out-Null)
        If($InstallationAttempts -gt 3) { break }
        $InstallationAttempts++
        Start-Sleep -Milliseconds 600
    }

    If(!($script:ReadyScript)) {
        return Write-Error "Unable to install ADAL Module: Please manually install the module by executing the command ""Install-Module -Name Microsoft.ADAL.PowerShell""."
    }
    Clear-Host

    try {
        <# Gather OAuth Bearer AccessToken from https://login.windows.net (Microsoft) #>
        Write-Host "[Get -> AccessToken]: Collecting AccessToken from Microsoft ~ Service -> ($script:ResourceId)...`n" -ForegroundColor Yellow
        Initialize-AccessToken -AuthorityName $script:AuthorityName -ClientId $script:AppId -ResourceId $script:ResourceId -RedirectUri $script:RedirectURI

        <# Connect to Exchange Online using the AccessToken gathered from https://login.windows.net (Microsoft) #>
        Write-Host "[Connect -> ExchangeOnline]: Aquired AccessToken - Connecting to Exchange Online`n" -ForegroundColor Yellow
        $SecuredToken = ConvertTo-SecureString $script:AuthorizationToken -AsPlainText -Force
        $UserCredential = New-Object System.Management.Automation.PSCredential($script:UserPrincipalName, $SecuredToken)
        $ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $script:ConnectionURI -Credential $UserCredential -Authentication Basic -AllowRedirection

        <# Finally - Import the Exchange PowerShell Session to allow the user to use Exchange Online cmdlets #>
        Import-Module (Import-PSSession $ExchSession -DisableNameChecking -AllowClobber) -Global
        Write-Host "`n[Session -> Imported] Successfully connected to Exchange Online" -ForegroundColor Green
    } catch {
        $_.Exception # Prints Error (If Any Exception)
    }
}

<#  Type:   Function-Void
#   Desc:   Uses the ADAL Module & Connects to https://login.windows.net - Gathering AccessToken & UPN | Original-Function: {Get-ADALAccessToken}
#   Usage:  Initialize-AccessToken -AuthorityName "portalms.onmicrosoft.com" -ClientId "a0c73c16-a7e3-4564-9a95-2bdf47383716" 
            -ResourceId "https://outlook.office365.com" -RedirectUri "urn:ietf:wg:oauth:2.0:oob" -ForcePromptSignIn $false
#   Ref:    Initialize-AccessToken | <While-loop> #>
Function Initialize-AccessToken {
    param(
        [parameter(Mandatory=$false)][string]$AuthorityName,
        [parameter(Mandatory=$true)][string]$ClientId,
        [parameter(Mandatory=$true)][string]$ResourceId,
        [parameter(Mandatory=$true)][string]$RedirectUri,
        [parameter(Mandatory=$false)][switch]$ForcePromptSignIn
    )
    <# Imports the AuthenticationContext Class from the ADAL Module .NET #>
    $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($script:EndpointURI)
    try {
        if($RedirectUri -ne '') {
            $NewRedirectUri = New-Object System.Uri -ArgumentList $RedirectUri
            if($ForcePromptSignIn) {
                $PromptBehavior = [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Always
            } else {
                $PromptBehavior = [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Auto
            }
            $Result = $Context.AcquireToken($ResourceId, $ClientId, $NewRedirectUri, $PromptBehavior)
        } else {
            $UserCredential = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential($UserName, $Password)
            $Result = $Context.AcquireToken($ResourceId, $ClientId, $UserCredential)
        }
    } catch [Microsoft.IdentityModel.Clients.ActiveDirectory.AdalException] {
        $_.Exception # Prints Error (If AdalException)
    }
    #Set AccessToken variables
    Set-AuthorizationToken -AccessToken $Result.AccessToken
    Set-UserPrincipalName -UserPrincipalName $Result.UserInfo.DisplayableId
}

<#  Type:   Is-Boolean
#   Desc:   Attempts to Install & Import the ADAL PowerShell Module
#   Usage:  Assert-ReadyADALModule
#   Ref:    Initialize-AccessToken | <While-loop> #>
Function Assert-ReadyADALModule {
    Try {

        <# RemoteSigned Execution Policy is required for Importing the Adal.PowerShell Module #>
        $RunWithPowerShell = (($script:args -like "*if((Get-ExecutionPolicy ) -ne 'AllSigned') { Set-ExecutionPolicy -Scope Process Bypass }; & *") -and 
                             ((Get-ExecutionPolicy -Scope "Process") -eq "Bypass"))

        If($RunWithPowerShell) {
            $PowerShellFilePath = $script:args.Substring(($script:args.LastIndexOf('&')) + 2)
            If($PowerShellFilePath.Length -gt 1) {
                Start-Process powershell -ArgumentList "-noexit ", "-noprofile ", "-command &$PowerShellFilePath"
                exit
            }
        }

        <# Automatically set the Execution to RemoteSigned to allow the module to be installed #>
        $CurrentExecutionPolicy = Get-ExecutionPolicy -List | Where-Object ExecutionPolicy -like "*RemoteSigned*"
        If(($null -eq $CurrentExecutionPolicy)) {
            Write-Host "[Set -> ExecutionPolicy]: Attempting to set CurrentUser's ExecutionPolicy to RemoteSigned..." -ForegroundColor Yellow
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
        }

        <# Microsoft.ADAL.PowerShell Module is required - Checks if the Module Exists | If not, attempts to Install the Module #>
        $ModulePath = Get-Module -ListAvailable -Name Microsoft.ADAL.PowerShell
        If($null -ne $ModulePath -and $ModulePath.Length -gt 0) {

            <# Import Microsoft.ADAL.PowerShell Module - Checks if the Module is Imported | If not, script fails and can't continue without manual user input #>
            Import-Module -Name Microsoft.ADAL.PowerShell
            $ModuleImported = Get-Module | Where-Object Name -like "*Adal.PowerShell*"
            If($null -ne $ModuleImported -and $ModuleImported.Length -gt 0) {
                <# Return $true - Module has been successfully improted and is ready to be used by the script #>
                Write-Host "[Connect -> ExchangeOnlineOAuth]: READY - Proceeding to request connection to $script:ResourceId" -ForegroundColor Green
                Set-ReadyScript -isReady $true; return $true
            }
        } Else {
            <# Attempt to install the Microsoft.ADAL.PowerShell module, needed to run this script #>
            Write-Host "[Install -> Module -> Microsoft.ADAL.PowerShell]: REQUIRED - Attempting to install module, requires user input..." -ForegroundColor Yellow
            Install-Module -Name Microsoft.ADAL.PowerShell -Scope CurrentUser -Force
        }
    } catch {
        $_.Exception  # Prints Error (If Any Exception)
    }
    return $false
}

<#  Type:   Set-String
#   Desc:   Sets local script to store AuthorizationToken | If $null - script fails
#   Usage:  AuthorizationToken -AccessToken "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6InBpVmxsb1FEU01Le..."
#   Ref:    Initialize-AccessToken #>
Function Set-AuthorizationToken {
    param([parameter(Mandatory=$true)][string] $AccessToken)
    $script:AuthorizationToken = ("Bearer " + $AccessToken)
}

<#  Type:   Set-String
#   Desc:   Sets local script to store UserPrincipalName | If $null - script fails
#   Usage:  Set-UserPrincipalName -UserPrincipalName "kai@portal.ms"
#   Ref:    Initialize-AccessToken #>
Function Set-UserPrincipalName {
    param([parameter(Mandatory=$true)][string] $UserPrincipalName)
    $script:UserPrincipalName = $UserPrincipalName
}

<#  Type:   Set-Boolean
#   Desc:   Sets local script to store ReadyScript boolean | If $false - stops script
#   Usage:  Set-ReadyScript -isReady $true
#   Ref:    Assert-ReadyADALModule #>
Function Set-ReadyScript {
    param([parameter(Mandatory=$true)][boolean] $isReady)
    $script:ReadyScript = $isReady
}

<#  Type:   Call-Function (Main)
#   Desc:   Calls the Main Function to start the entire PowerShell script
#   Usage:  Connect-ExchangeOnlineOAuth | "Right-click -> Run with PowerShell on .ps1 file"
#   Ref:    NONE #>
Connect-ExchangeOnlineOAuth
