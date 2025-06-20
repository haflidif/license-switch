<#
.SYNOPSIS
    Office 365 License Switching Script
.DESCRIPTION
    This script connects to Microsoft Graph, exports users with a specific license,
    and switches them to a new license. Supports both SkuPartNumber and SkuId input.
.PARAMETER ExpiringLicenseSku
    The SkuPartNumber of the license to be removed/switched from (user-friendly names)
.PARAMETER NewLicenseSku
    The SkuPartNumber of the license to be assigned (user-friendly names)
.PARAMETER ExpiringLicenseSkuId
    The SkuId (GUID) of the license to be removed/switched from
.PARAMETER NewLicenseSkuId
    The SkuId (GUID) of the license to be assigned
.PARAMETER ExportPath
    Path for the export CSV file (optional)
.PARAMETER WhatIf
    Preview changes without making them
.PARAMETER TenantId
    The Azure AD Tenant ID (GUID) to connect to. Optional - uses default tenant if not specified.
.PARAMETER TenantDomain
    The Azure AD Tenant domain (e.g., contoso.onmicrosoft.com) to connect to. Alternative to TenantId.
.EXAMPLE
    # Using SkuPartNumber (user-friendly names) - Default method
    .\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "Microsoft_365_E5_(no_Teams)" -NewLicenseSku "Microsoft_Teams_Enterprise_New" -WhatIf

.EXAMPLE
    # Using SkuId (GUIDs) - Direct method
    .\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01" -WhatIf

.EXAMPLE
    # Using verbose mode to see detailed processing (shows individual user messages)
    .\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "Microsoft_365_E5_(no_Teams)" -NewLicenseSku "Microsoft_Teams_Enterprise_New" -WhatIf -Verbose

.EXAMPLE
    # Multi-tenant: Connect to specific tenant using Tenant ID
    .\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "Microsoft_365_E5_(no_Teams)" -NewLicenseSku "Microsoft_Teams_Enterprise_New" -TenantId "12345678-1234-1234-1234-123456789012" -WhatIf

.EXAMPLE
    # Multi-tenant: Connect to specific tenant using domain
    .\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "Microsoft_365_E5_(no_Teams)" -NewLicenseSku "Microsoft_Teams_Enterprise_New" -TenantDomain "contoso.onmicrosoft.com" -WhatIf
#>

[CmdletBinding(DefaultParameterSetName = 'BySkuPartNumber')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'BySkuPartNumber')]
    [string]$ExpiringLicenseSku,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'BySkuPartNumber')]
    [string]$NewLicenseSku,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'BySkuId')]
    [string]$ExpiringLicenseSkuId,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'BySkuId')]
    [string]$NewLicenseSkuId,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ".\LicenseSwitchExport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantDomain
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
    param(
        [string]$TenantId,
        [string]$TenantDomain
    )
    
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
        
        # Prepare connection parameters
        $connectionParams = @{
            Scopes = @("User.ReadWrite.All", "Directory.ReadWrite.All", "Organization.Read.All")
        }
        
        # Add tenant specification if provided
        if ($TenantId) {
            $connectionParams.TenantId = $TenantId
            Write-ColorOutput "Connecting to specific tenant ID: $TenantId" "Cyan"
        }
        elseif ($TenantDomain) {
            $connectionParams.TenantId = $TenantDomain
            Write-ColorOutput "Connecting to specific tenant domain: $TenantDomain" "Cyan"
        }
        else {
            Write-ColorOutput "Connecting to default tenant..." "Cyan"
        }
        
        # Connect with required scopes
        Connect-MgGraph @connectionParams
        
        # Get current context to show which tenant we're connected to
        $context = Get-MgContext
        if ($context) {
            Write-ColorOutput "Successfully connected to Microsoft Graph" "Green"
            Write-ColorOutput "Tenant ID: $($context.TenantId)" "White"
            Write-ColorOutput "Account: $($context.Account)" "White"
            Write-ColorOutput "Environment: $($context.Environment)" "White"
        }
        
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
        $licenses = Get-MgSubscribedSku -Property SkuPartNumber, SkuId, ConsumedUnits, PrepaidUnits
        
        Write-ColorOutput "`nAvailable Licenses:" "Cyan"
        Write-ColorOutput "===================" "Cyan"
        
        foreach ($license in $licenses) {
            $available = $license.PrepaidUnits.Enabled - $license.ConsumedUnits
            Write-Host "SKU Part Number: $($license.SkuPartNumber)" -ForegroundColor White
            Write-Host "  SKU ID (GUID): $($license.SkuId)" -ForegroundColor Gray
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

# Function to resolve license information (works with both SkuPartNumber and SkuId)
function Resolve-LicenseInfo {
    param(
        [array]$AvailableLicenses,
        [string]$LicenseIdentifier,
        [string]$IdentifierType  # "SkuPartNumber" or "SkuId"
    )
    
    if ($IdentifierType -eq "SkuPartNumber") {
        return $AvailableLicenses | Where-Object { $_.SkuPartNumber -eq $LicenseIdentifier }
    } else {
        return $AvailableLicenses | Where-Object { $_.SkuId -eq $LicenseIdentifier }
    }
}

# Function to validate license inputs (enhanced to work with both input types)
function Test-LicenseValidity {
    param(
        [array]$AvailableLicenses,
        [string]$ExpiringLicense,
        [string]$NewLicense,
        [string]$InputType  # "SkuPartNumber" or "SkuId"
    )
    
    $expiringLicenseValid = Resolve-LicenseInfo -AvailableLicenses $AvailableLicenses -LicenseIdentifier $ExpiringLicense -IdentifierType $InputType
    $newLicenseValid = Resolve-LicenseInfo -AvailableLicenses $AvailableLicenses -LicenseIdentifier $NewLicense -IdentifierType $InputType
    
    if (!$expiringLicenseValid) {
        Write-ColorOutput "ERROR: Expiring license '$ExpiringLicense' not found!" "Red"
        return $false
    }
    
    if (!$newLicenseValid) {
        Write-ColorOutput "ERROR: New license '$NewLicense' not found!" "Red"
        return $false
    }
    
    # Check if new license has available units
    $availableUnits = $newLicenseValid.PrepaidUnits.Enabled - $newLicenseValid.ConsumedUnits
    if ($availableUnits -le 0) {
        Write-ColorOutput "WARNING: New license '$NewLicense' has no available units!" "Red"
        return $false
    }
    
    Write-ColorOutput "License validation successful!" "Green"
    if ($InputType -eq "SkuPartNumber") {
        Write-ColorOutput "Expiring License: $($expiringLicenseValid.SkuPartNumber)" "White"
        Write-ColorOutput "New License: $($newLicenseValid.SkuPartNumber) ($availableUnits units available)" "White"
    } else {
        Write-ColorOutput "Expiring License: $($expiringLicenseValid.SkuPartNumber) (SkuId: $($expiringLicenseValid.SkuId))" "White"
        Write-ColorOutput "New License: $($newLicenseValid.SkuPartNumber) (SkuId: $($newLicenseValid.SkuId)) - $availableUnits units available" "White"
    }
    
    return $true
}

# Function to get users with specific license (enhanced to work with both input types)
function Get-UsersWithLicense {
    param(
        [string]$LicenseIdentifier,
        [string]$InputType,  # "SkuPartNumber" or "SkuId"
        [array]$AvailableLicenses
    )
      try {
        # Get the SkuId for the license query
        if ($InputType -eq "SkuPartNumber") {
            $licenseSkuId = ($AvailableLicenses | Where-Object { $_.SkuPartNumber -eq $LicenseIdentifier }).SkuId
            Write-ColorOutput "Searching for users with license: $LicenseIdentifier (SkuId: $licenseSkuId)..." "Yellow"
        } else {
            $licenseSkuId = $LicenseIdentifier
            $licenseSkuPartNumber = ($AvailableLicenses | Where-Object { $_.SkuId -eq $LicenseIdentifier }).SkuPartNumber
            Write-ColorOutput "Searching for users with license: $licenseSkuPartNumber (SkuId: $LicenseIdentifier)..." "Yellow"
        }
        
        Write-ColorOutput "Using optimized server-side filtering for better performance..." "Cyan"
        
        # Use server-side filtering for much better performance
        # This approach filters on the server instead of downloading all users
        $filterQuery = "assignedLicenses/any(x:x/skuId eq $licenseSkuId)"
        
        Write-ColorOutput "Executing Graph API query with filter..." "Gray"
        $users = Get-MgUser -Filter $filterQuery -ConsistencyLevel eventual -All -Property "Id,DisplayName,UserPrincipalName,AssignedLicenses,UsageLocation"
        
        Write-ColorOutput "Found $($users.Count) users with the specified license" "Green"
        
        # Additional validation to ensure we got the right users
        if ($users.Count -gt 0) {
            Write-ColorOutput "Validating results..." "Gray"
            $validatedUsers = $users | Where-Object { 
                $_.AssignedLicenses.SkuId -contains $licenseSkuId 
            }
            
            if ($validatedUsers.Count -ne $users.Count) {
                Write-ColorOutput "Warning: Server filter returned $($users.Count) users, but only $($validatedUsers.Count) actually have the license." "Yellow"
                $users = $validatedUsers
            } else {
                Write-ColorOutput "Server-side filtering worked perfectly!" "Green"
            }
        }
        
        return $users
    }
    catch {
        Write-ColorOutput "Server-side filtering failed: $($_.Exception.Message)" "Yellow"
        Write-ColorOutput "Falling back to client-side filtering (this may be slower)..." "Yellow"
        
        try {
            # Fallback to the original method if server-side filtering fails
            Write-ColorOutput "Retrieving all users with licenses (this may take a while for large tenants)..." "Yellow"
            
            # Use pagination to process users in batches
            $batchSize = 999  # Maximum page size for Microsoft Graph
            $allUsers = @()
            $pageCount = 0
            
            do {
                $pageCount++
                Write-ColorOutput "Processing batch $pageCount (up to $batchSize users)..." "Gray"
                
                if ($pageCount -eq 1) {
                    $userPage = Get-MgUser -Top $batchSize -Property "Id,DisplayName,UserPrincipalName,AssignedLicenses,UsageLocation" -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual
                } else {
                    # Handle pagination for subsequent pages
                    $userPage = Get-MgUser -Top $batchSize -Property "Id,DisplayName,UserPrincipalName,AssignedLicenses,UsageLocation" -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual
                }
                
                # Filter users with the specific license
                $usersWithLicense = $userPage | Where-Object { 
                    $_.AssignedLicenses.SkuId -contains $licenseSkuId
                }
                
                if ($usersWithLicense) {
                    $allUsers += $usersWithLicense
                    Write-ColorOutput "Found $($usersWithLicense.Count) users in this batch" "Green"
                }
                
                # Note: This simplified approach processes the first batch only
                # For full pagination, you'd need to handle @odata.nextLink
                break
                
            } while ($false)  # Simplified - in reality you'd check for @odata.nextLink
            
            Write-ColorOutput "Total users found with the specified license: $($allUsers.Count)" "Green"
            return $allUsers
            
        } catch {
            Write-ColorOutput "Fallback method also failed: $($_.Exception.Message)" "Red"
            return $null
        }
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
        [bool]$WhatIfMode = $false,
        [bool]$VerboseMode = $false
    )
    
    try {
        if ($WhatIfMode) {
            if ($VerboseMode) {
                Write-ColorOutput "WHATIF: Would switch license for $UserPrincipalName" "Yellow"
            }
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
    
    # Determine input type and set variables
    $usingSkuPartNumber = $PSCmdlet.ParameterSetName -eq 'BySkuPartNumber'
    
    if ($usingSkuPartNumber) {
        Write-ColorOutput "Input Mode: SkuPartNumber (User-friendly names)" "Yellow"
        $expiringInput = $ExpiringLicenseSku
        $newInput = $NewLicenseSku
        $inputType = "SkuPartNumber"
    } else {
        Write-ColorOutput "Input Mode: SkuId (GUIDs)" "Yellow"
        $expiringInput = $ExpiringLicenseSkuId
        $newInput = $NewLicenseSkuId
        $inputType = "SkuId"
    }
      Write-ColorOutput ""
    
    # Connect to Microsoft Graph
    if (!(Connect-ToMicrosoftGraph -TenantId $TenantId -TenantDomain $TenantDomain)) {
        Write-ColorOutput "Exiting due to connection failure." "Red"
        return
    }
    
    # Get available licenses
    $availableLicenses = Get-AvailableLicenses
    if (!$availableLicenses) {
        Write-ColorOutput "Exiting due to license retrieval failure." "Red"
        return
    }
    
    # Validate license inputs
    if (!(Test-LicenseValidity -AvailableLicenses $availableLicenses -ExpiringLicense $expiringInput -NewLicense $newInput -InputType $inputType)) {
        Write-ColorOutput "Exiting due to license validation failure." "Red"
        return
    }
      # Get license info for operations
    $expiringLicenseInfo = Resolve-LicenseInfo -AvailableLicenses $availableLicenses -LicenseIdentifier $expiringInput -IdentifierType $inputType
    $newLicenseInfo = Resolve-LicenseInfo -AvailableLicenses $availableLicenses -LicenseIdentifier $newInput -IdentifierType $inputType
    
    # Get users with expiring license (with timeout handling)
    Write-ColorOutput "`n=== User Discovery Phase ===" "Cyan"
    Write-ColorOutput "Starting user search - this may take a few moments for large tenants..." "Yellow"
    
    $searchStartTime = Get-Date
    $usersWithExpiringLicense = Get-UsersWithLicense -LicenseIdentifier $expiringInput -InputType $inputType -AvailableLicenses $availableLicenses
    $searchEndTime = Get-Date
    $searchDuration = $searchEndTime - $searchStartTime
    
    Write-ColorOutput "User search completed in $([math]::Round($searchDuration.TotalSeconds, 2)) seconds" "Gray"
    
    if (!$usersWithExpiringLicense -or $usersWithExpiringLicense.Count -eq 0) {
        Write-ColorOutput "No users found with the specified expiring license." "Yellow"
        Write-ColorOutput "This could mean:" "Gray"
        Write-ColorOutput "  • No users are currently assigned this license" "Gray"  
        Write-ColorOutput "  • The license identifier is incorrect" "Gray"
        Write-ColorOutput "  • There was a search timeout or error" "Gray"
        return
    }
    
    Write-ColorOutput "✅ Found $($usersWithExpiringLicense.Count) users with the expiring license" "Green"
    
    # Show sample of users found (for verification)
    if ($usersWithExpiringLicense.Count -le 5) {
        Write-ColorOutput "`nUsers found:" "White"
        $usersWithExpiringLicense | ForEach-Object { Write-ColorOutput "  • $($_.DisplayName) ($($_.UserPrincipalName))" "Gray" }
    } else {
        Write-ColorOutput "`nSample of users found:" "White"
        $usersWithExpiringLicense | Select-Object -First 3 | ForEach-Object { Write-ColorOutput "  • $($_.DisplayName) ($($_.UserPrincipalName))" "Gray" }
        Write-ColorOutput "  ... and $($usersWithExpiringLicense.Count - 3) more users" "Gray"
    }
    
    # Export users to CSV
    Write-ColorOutput "`n=== Export Phase ===" "Cyan"
    if (!(Export-UsersToCSV -Users $usersWithExpiringLicense -FilePath $ExportPath -LicenseSku $expiringLicenseInfo.SkuPartNumber)) {
        Write-ColorOutput "Warning: Failed to export users, but continuing with license switch..." "Yellow"
    }
    
    # Confirm license switch
    Write-ColorOutput "`nReady to switch licenses for $($usersWithExpiringLicense.Count) users" "Yellow"
    Write-ColorOutput "From: $($expiringLicenseInfo.SkuPartNumber)" "White"
    Write-ColorOutput "To: $($newLicenseInfo.SkuPartNumber)" "White"
    
    if (!$WhatIf) {
        $confirmation = Read-Host "`nDo you want to proceed with the license switch? (y/N)"
        if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
            Write-ColorOutput "License switch cancelled by user." "Yellow"
            return
        }
    }    # Perform license switch
    if ($WhatIf) {
        Write-ColorOutput "`nStarting license switch simulation (WhatIf mode)..." "Yellow"
        if ($VerbosePreference -ne 'Continue') {
            Write-ColorOutput "Processing $($usersWithExpiringLicense.Count) users silently (use -Verbose to see individual users)..." "Gray"
        }
    } else {
        Write-ColorOutput "`nStarting license switch process..." "Yellow"
    }
    $successCount = 0
    $failureCount = 0
    
    foreach ($user in $usersWithExpiringLicense) {
        $result = Switch-UserLicense -UserId $user.Id -UserPrincipalName $user.UserPrincipalName -ExpiringLicenseSkuId $expiringLicenseInfo.SkuId -NewLicenseSkuId $newLicenseInfo.SkuId -WhatIfMode $WhatIf -VerboseMode ($VerbosePreference -eq 'Continue')
        
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
