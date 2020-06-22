function Connect-Office365 {
    [CmdletBinding(DefaultParameterSetName = "UseSavedDefaultCredential")]
    param(
        [Parameter(Mandatory = $false, ParameterSetName = "UseSavedDefaultCredential" )]
        [bool]
        $UseSavedDefaultCredential = $true,

        [Parameter(Mandatory = $true, ParameterSetName = "UseSavedCredential" )]
        [string]
        $SavedUserName,

        [Parameter(Mandatory = $true, ParameterSetName = "UseManualCredential" )]
        [switch]
        $UseManualCredential,

        [Parameter(Mandatory = $false, ParameterSetName = "UseSavedDefaultCredential" )]
        [Parameter(Mandatory = $false, ParameterSetName = "UseSavedCredential" )]
        [Parameter(Mandatory = $false, ParameterSetName = "UseManualCredential" )]
        [ValidateSet("MSOnline", "AzureAD", "EXO", "SCC")]
        [string[]]
        $Services = @("MSOnline", "EXO")
    )

    Import-Module ExchangeOnlineManagement

    $ConfigPath = Join-Path -Path:([Environment]::GetFolderPath('MyDocuments')) -ChildPath "Office365Credential.xml"

    [pscredential]$Credential = $null

    if ($PSCmdlet.ParameterSetName -eq "UseSavedDefaultCredential") {
        $DefaultUser = (Import-Clixml $ConfigPath) | Where-Object {$_.IsDefault -eq $true}

        if ($DefaultUser) {
            $Credential = $DefaultUser.Credential
        }
        else {
            Write-Error "You don't have the saved default credential."
            return
        }
    }
    elseif ($PSCmdlet.ParameterSetName -eq "UseSavedCredential") {
        $SavedUser = (Import-Clixml $ConfigPath) | Where-Object {$_.Name -eq $SavedUserName}

        if ($SavedUser) {
            $Credential = $SavedUser.Credential
        }
        else {
            Write-Error "The specified user was not found."
        }
    }
    else {
        try {
            $Credential = Get-Credential
        }
        catch {
            return
        }
    }

    Write-Host "Connecting to Office 365 using $($Credential.UserName)"

    try {
        Disconnect-AzureAD 2>&1 | Out-Null
    }
    catch {
    }

    # Disconnect-ExchangeOnline -Confirm:$false *>&1 | Out-Null

    if ($Services -contains "EXO") {
        Write-Host "Connecting to Exchange Online"

        Connect-ExchangeOnline -Credential $Credential -ShowProgress $true -ShowBanner:$false
    }

    if ($Services -contains "MSOnline") {
        Write-Host "Connecting to Azure Active Directory (MSOnline)"

        Connect-MsolService -Credential $Credential
    }

    if ($Services -contains "AzureAD") {
        Write-Host "Connecting to Azure Active Directory (AzureAD)"

        Connect-AzureAD -Credential $Credential
    }

    if ($Services -contains "SCC") {
        if ($Services -contains "EXO") {
            Write-Warning "You cannot connect to SCC and EXO at the same time."
        }
        else{
            Write-Host "Connecting to Security & Compliance Center"
            Connect-IPPSSession -Credential $Credential
        }
    }

    # if ($Services -contains "MSOnline") {
    #     Write-Host "Tenant Expiration Information"

    #     $Now = (Get-Date).ToUniversalTime()
    #     $SubscriptionInfo = Get-MsolSubscription | Select-Object SkuPartNumber, NextLifecycleDate

    #     foreach ($Subscription in $SubscriptionInfo) {
    #         $DaysToExpiration = ($Subscription.NextLifecycleDate - $Now).TotalDays

    #         if ($DaysToExpiration -le 30) {
    #             if ($DaysToExpiration -ge 1) {
    #                 Write-Host "$($Subscription.SkuPartNumber) will be expired in $DaysToExpiration days."
    #             }
    #             else {
    #                 Write-Host "$($Subscription.SkuPartNumber) is expired."
    #             }
    #         }
    #     }
    # }

    if ($Services -contains "EXO" -or $Services -contains "SCC") {
        $ModuleVersion = (Get-Module ExchangeOnlineManagement).Version
        Write-Host "You are using ExchangeOnlineManagement version $($ModuleVersion.ToString()) to connect Exchange Online or Security Compliance Center. If you want to check if the newer module is available, run 'Test-ExchangeOnlineManagementModuleVersion' cmdlet."
    }
}

function New-Office365Credential {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $false)]
        [pscredential]$Credential,
        [Parameter(Mandatory = $false, Position = 2, ValueFromPipeline = $false)]
        [string]$Name = "",
        [Parameter(Mandatory = $false, Position = 3, ValueFromPipeline = $false)]
        [bool]$IsDefault = $false
    )

    if ($Name -eq "") {
        $Name = $Credential.UserName
    }

    $ConfigPath = Join-Path -Path:([Environment]::GetFolderPath('MyDocuments')) -ChildPath "Office365Credential.xml"

    if (Test-Path $ConfigPath) {
        $CurrentConfig = Import-Clixml $ConfigPath

        if (($CurrentConfig.Name) -contains $Name) {
            Write-Error "$Name is already registered."
            return
        }

        if ($IsDefault) {
            $CurrentConfig | ForEach-Object {$_.IsDefault = $false}
        }

        $CurrentConfig.Add(
            [PSCustomObject]@{
                Name       = $Name
                Credential = $Credential
                IsDefault  = $IsDefault
            }
        ) | Out-Null

        Export-Clixml -InputObject $CurrentConfig -Path $ConfigPath
    }
    else {
        $NewConfig = New-Object 'System.Collections.Generic.List[PSCustomObject]'

        $NewConfig.Add([PSCustomObject]@{
                Name       = $Name
                Credential = $Credential
                IsDefault  = $true
            }) | Out-Null

        Export-Clixml -InputObject $NewConfig -Path $ConfigPath
    }
}

function Get-Office365Credential {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false, Position = 1, ValueFromPipeline = $true)]
        [string]$Name
    )

    $ConfigPath = Join-Path -Path:([Environment]::GetFolderPath('MyDocuments')) -ChildPath "Office365Credential.xml"
    $Result = $null

    if (Test-Path $ConfigPath) {
        $CurrentConfig = Import-Clixml $ConfigPath

        if ($null -ne $Name -and "" -ne $Name) {
            $Result = $CurrentConfig | Where-Object {$_.Name -like $Name}
        }
        else {
            $Result = $CurrentConfig
        }
    }
    else {
        $Result = $null
    }

    if ($null -eq $Result) {
        throw [System.Exception] "$Name is not found."
    }
    else {
        return $Result
    }
}

function Remove-Office365Credential {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
    param
    (
        [Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
        [string]$Name
    )

    $CredentialNameToBeRemoved = Get-Office365Credential -Name $Name | Select-Object -ExpandProperty Name

    if ($null -eq $CredentialNameToBeRemoved) {
        return
    }

    $CurrentConfig = Get-Office365Credential

    foreach ($Item in $CredentialNameToBeRemoved) {
        if ($PSCmdlet.ShouldProcess($Item, "Remove-Office365Credential")) {
            $CurrentConfig = $CurrentConfig | Where-Object {$CredentialNameToBeRemoved -notcontains $_.Name}
        }
    }
    
    $ConfigPath = Join-Path -Path:([Environment]::GetFolderPath('MyDocuments')) -ChildPath "Office365Credential.xml"

    Export-Clixml -InputObject $CurrentConfig -Path $ConfigPath
}

function Copy-ConnectExoCommand {
	'$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential:(Get-Credential) -Authentication Basic -AllowRedirection; Import-PSSession $Session -DisableNameChecking' | clip
}

function Test-ExchangeOnlineManagementModuleVersion {
    Write-Progress -Activity "Checking if a newer version of ExchangeOnlineManagement module is available." -Status "Loading installed modules." -PercentComplete 10

    $InstalledModules = Get-Module ExchangeOnlineManagement -ListAvailable

    $InstalledVersion = New-Object System.Version(0, 0, 0, 0)

    foreach ($Module in $InstalledModules) {
        if ($Module.Version -gt $InstalledVersion) {
            $InstalledVersion = $Module.Version
        }
    }

    Write-Progress -Activity "Checking if a newer version of ExchangeOnlineManagement module is available..." -Status "Running 'Find-Module ExchangeOnlineManagement" -PercentComplete 20

    $LatestModules = Find-Module ExchangeOnlineManagement
    $LatestVersion = New-Object System.Version(0, 0, 0, 0)

    foreach ($Module in $LatestModules) {
        if ($Module.Version -gt $LatestVersion) {
            $LatestVersion = $Module.Version
        }
    }

    Write-Progress -Activity "Checking if a newer version of ExchangeOnlineManagement module is available." -Completed

    if ($LatestVersion -gt $InstalledVersion) {
        Write-Host "New ExchangeOnlineManagement module is available. Run 'Uninstall-Module ExchangeOnlineManagement -AllVersions; Install-Module -Name ExchangeOnlineManagement' from an elevated Windows PowerShell window."
        $LatestModules | Format-Table -AutoSize
    }
    else {
        Write-Host "You are using the latest ExchangeOnlineManagement module."
    }
}