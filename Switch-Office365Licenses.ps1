<#
.SYNOPSIS
    Office 365 License Switching Script
.DESCRIPTION
    This script connects to Microsoft Graph, exports users with a specific license,
    and switches them to a new license. It includes validation and logging.
.PARAMETER ExpiringLicenseSku
    The SKU of the license to be removed/switched from
.PARAMETER NewLicenseSku
    The SKU of the license to be assigned
.PARAMETER ExportPath
    Path for the export CSV file (optional)
.PARAMETER WhatIf
    Preview changes without making them
.EXAMPLE
    .\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3" -WhatIf
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExpiringLicenseSku,
    
    [Parameter(Mandatory = $true)]
    [string]$NewLicenseSku,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ".\LicenseSwitchExport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    try {
        Write-ColorOutput "Connecting to Microsoft Graph..." "Yellow"
        
        # Check if Microsoft.Graph module is installed
        if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
            Write-ColorOutput "Microsoft.Graph module not found. Installing..." "Yellow"
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
        }
        
        # Import required modules
        Import-Module Microsoft.Graph.Authentication
        Import-Module Microsoft.Graph.Users
        Import-Module Microsoft.Graph.Identity.DirectoryManagement
        
        # Connect with required scopes
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "Organization.Read.All"
        
        Write-ColorOutput "Successfully connected to Microsoft Graph" "Green"
        return $true
    }
    catch {
        Write-ColorOutput "Failed to connect to Microsoft Graph: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to get all available licenses
function Get-AvailableLicenses {
    try {
        Write-ColorOutput "Retrieving available licenses..." "Yellow"
        $licenses = Get-MgSubscribedSku
        
        Write-ColorOutput "`nAvailable Licenses:" "Cyan"
        Write-ColorOutput "===================" "Cyan"
        
        foreach ($license in $licenses) {
            $available = $license.PrepaidUnits.Enabled - $license.ConsumedUnits
            Write-Host "SKU: $($license.SkuPartNumber)" -ForegroundColor White
            Write-Host "  Display Name: $($license.SkuId)" -ForegroundColor Gray
            Write-Host "  Total Units: $($license.PrepaidUnits.Enabled)" -ForegroundColor Gray
            Write-Host "  Consumed: $($license.ConsumedUnits)" -ForegroundColor Gray
            Write-Host "  Available: $available" -ForegroundColor $(if($available -gt 0) {"Green"} else {"Red"})
            Write-Host ""
        }
        
        return $licenses
    }
    catch {
        Write-ColorOutput "Failed to retrieve licenses: $($_.Exception.Message)" "Red"
        return $null
    }
}

# Function to validate license SKUs
function Test-LicenseValidity {
    param(
        [array]$AvailableLicenses,
        [string]$ExpiringLicense,
        [string]$NewLicense
    )
    
    $expiringLicenseValid = $AvailableLicenses | Where-Object { $_.SkuPartNumber -eq $ExpiringLicense }
    $newLicenseValid = $AvailableLicenses | Where-Object { $_.SkuPartNumber -eq $NewLicense }
    
    if (!$expiringLicenseValid) {
        Write-ColorOutput "ERROR: Expiring license SKU '$ExpiringLicense' not found!" "Red"
        return $false
    }
    
    if (!$newLicenseValid) {
        Write-ColorOutput "ERROR: New license SKU '$NewLicense' not found!" "Red"
        return $false
    }
    
    # Check if new license has available units
    $availableUnits = $newLicenseValid.PrepaidUnits.Enabled - $newLicenseValid.ConsumedUnits
    if ($availableUnits -le 0) {
        Write-ColorOutput "WARNING: New license '$NewLicense' has no available units!" "Red"
        return $false
    }
    
    Write-ColorOutput "License validation successful!" "Green"
    Write-ColorOutput "Expiring License: $($expiringLicenseValid.SkuPartNumber)" "White"
    Write-ColorOutput "New License: $($newLicenseValid.SkuPartNumber) ($availableUnits units available)" "White"
    
    return $true
}

# Function to get users with specific license
function Get-UsersWithLicense {
    param(
        [string]$LicenseSku
    )
    
    try {
        Write-ColorOutput "Searching for users with license: $LicenseSku..." "Yellow"
        
        # Get all users with licenses
        $users = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,AssignedLicenses" | 
                 Where-Object { 
                     $_.AssignedLicenses.SkuId -contains (Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $LicenseSku }).SkuId 
                 }
        
        Write-ColorOutput "Found $($users.Count) users with license: $LicenseSku" "Green"
        
        return $users
    }
    catch {
        Write-ColorOutput "Failed to retrieve users: $($_.Exception.Message)" "Red"
        return $null
    }
}

# Function to export users to CSV
function Export-UsersToCSV {
    param(
        [array]$Users,
        [string]$FilePath,
        [string]$LicenseSku
    )
    
    try {
        $exportData = $Users | Select-Object @{
            Name = 'DisplayName'
            Expression = { $_.DisplayName }
        }, @{
            Name = 'UserPrincipalName'
            Expression = { $_.UserPrincipalName }
        }, @{
            Name = 'UserId'
            Expression = { $_.Id }
        }, @{
            Name = 'CurrentLicense'
            Expression = { $LicenseSku }
        }, @{
            Name = 'ExportDate'
            Expression = { Get-Date -Format 'yyyy-MM-dd HH:mm:ss' }
        }
        
        $exportData | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        Write-ColorOutput "Users exported to: $FilePath" "Green"
        
        return $true
    }
    catch {
        Write-ColorOutput "Failed to export users: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to switch user license
function Switch-UserLicense {
    param(
        [string]$UserId,
        [string]$UserPrincipalName,
        [string]$ExpiringLicenseSkuId,
        [string]$NewLicenseSkuId,
        [bool]$WhatIfMode = $false
    )
    
    try {
        if ($WhatIfMode) {
            Write-ColorOutput "WHATIF: Would switch license for $UserPrincipalName" "Yellow"
            return $true
        }
        
        # Remove old license and add new license
        $licenseParams = @{
            AddLicenses = @(
                @{
                    SkuId = $NewLicenseSkuId
                }
            )
            RemoveLicenses = @($ExpiringLicenseSkuId)
        }
        
        Set-MgUserLicense -UserId $UserId -BodyParameter $licenseParams
        Write-ColorOutput "Successfully switched license for: $UserPrincipalName" "Green"
        
        return $true
    }
    catch {
        Write-ColorOutput "Failed to switch license for $UserPrincipalName`: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Main execution
function Main {
    Write-ColorOutput "=== Office 365 License Switching Script ===" "Cyan"
    Write-ColorOutput "Start Time: $(Get-Date)" "Gray"
    Write-ColorOutput ""
    
    # Connect to Microsoft Graph
    if (!(Connect-ToMicrosoftGraph)) {
        Write-ColorOutput "Exiting due to connection failure." "Red"
        return
    }
    
    # Get available licenses
    $availableLicenses = Get-AvailableLicenses
    if (!$availableLicenses) {
        Write-ColorOutput "Exiting due to license retrieval failure." "Red"
        return
    }
    
    # Validate license SKUs
    if (!(Test-LicenseValidity -AvailableLicenses $availableLicenses -ExpiringLicense $ExpiringLicenseSku -NewLicense $NewLicenseSku)) {
        Write-ColorOutput "Exiting due to license validation failure." "Red"
        return
    }
    
    # Get SKU IDs
    $expiringLicenseSkuId = ($availableLicenses | Where-Object { $_.SkuPartNumber -eq $ExpiringLicenseSku }).SkuId
    $newLicenseSkuId = ($availableLicenses | Where-Object { $_.SkuPartNumber -eq $NewLicenseSku }).SkuId
    
    # Get users with expiring license
    $usersWithExpiringLicense = Get-UsersWithLicense -LicenseSku $ExpiringLicenseSku
    if (!$usersWithExpiringLicense -or $usersWithExpiringLicense.Count -eq 0) {
        Write-ColorOutput "No users found with the specified expiring license." "Yellow"
        return
    }
    
    # Export users to CSV
    if (!(Export-UsersToCSV -Users $usersWithExpiringLicense -FilePath $ExportPath -LicenseSku $ExpiringLicenseSku)) {
        Write-ColorOutput "Warning: Failed to export users, but continuing with license switch..." "Yellow"
    }
    
    # Confirm license switch
    Write-ColorOutput "`nReady to switch licenses for $($usersWithExpiringLicense.Count) users" "Yellow"
    Write-ColorOutput "From: $ExpiringLicenseSku" "White"
    Write-ColorOutput "To: $NewLicenseSku" "White"
    
    if (!$WhatIf) {
        $confirmation = Read-Host "`nDo you want to proceed with the license switch? (y/N)"
        if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
            Write-ColorOutput "License switch cancelled by user." "Yellow"
            return
        }
    }
    
    # Perform license switch
    Write-ColorOutput "`nStarting license switch process..." "Yellow"
    $successCount = 0
    $failureCount = 0
    
    foreach ($user in $usersWithExpiringLicense) {
        $result = Switch-UserLicense -UserId $user.Id -UserPrincipalName $user.UserPrincipalName -ExpiringLicenseSkuId $expiringLicenseSkuId -NewLicenseSkuId $newLicenseSkuId -WhatIfMode $WhatIf
        
        if ($result) {
            $successCount++
        } else {
            $failureCount++
        }
        
        # Add a small delay to avoid throttling
        Start-Sleep -Milliseconds 100
    }
    
    # Summary
    Write-ColorOutput "`n=== License Switch Summary ===" "Cyan"
    Write-ColorOutput "Total Users Processed: $($usersWithExpiringLicense.Count)" "White"
    Write-ColorOutput "Successful: $successCount" "Green"
    Write-ColorOutput "Failed: $failureCount" "Red"
    Write-ColorOutput "Export File: $ExportPath" "White"
    Write-ColorOutput "End Time: $(Get-Date)" "Gray"
    
    if ($WhatIf) {
        Write-ColorOutput "`nNote: This was a WHATIF run. No actual changes were made." "Yellow"
    }
}

# Execute main function
Main
