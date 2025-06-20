<#
.SYNOPSIS
    Quick License Lookup Script
.DESCRIPTION
    This helper script connects to Microsoft Graph and displays all available
    Office 365 licenses with their SKUs, making it easier to identify the
    correct license names for the main switching script.
.PARAMETER TenantId
    The Azure AD Tenant ID (GUID) to connect to. Optional - if not provided, will use the default tenant.
.PARAMETER TenantDomain
    The Azure AD Tenant domain (e.g., contoso.onmicrosoft.com) to connect to. Alternative to TenantId.
.EXAMPLE
    .\Get-Office365Licenses.ps1
.EXAMPLE
    .\Get-Office365Licenses.ps1 -TenantId "12345678-1234-1234-1234-123456789012"
.EXAMPLE
    .\Get-Office365Licenses.ps1 -TenantDomain "contoso.onmicrosoft.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9]*\.onmicrosoft\.com$|^[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9]*\.[a-zA-Z]{2,}$')]
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
        Import-Module Microsoft.Graph.Identity.DirectoryManagement
        
        # Prepare connection parameters
        $connectionParams = @{
            Scopes = @("Directory.Read.All", "Organization.Read.All")
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

# Function to get and display all available licenses
function Get-AvailableLicenses {
    try {
        Write-ColorOutput "Retrieving available licenses..." "Yellow"
        
        # Use the recommended Microsoft Graph PowerShell command
        $licenses = Get-MgSubscribedSku -Property SkuPartNumber, SkuId, ConsumedUnits, PrepaidUnits | Sort-Object SkuPartNumber
          Write-ColorOutput "`n=== Available Office 365 Licenses in Your Tenant ===" "Cyan"
        Write-ColorOutput "=====================================================" "Cyan"
        Write-ColorOutput "Note: This shows only licenses your organization has purchased/subscribed to." "Yellow"
        Write-ColorOutput "      It does not show all possible Microsoft 365 licenses that exist." "Yellow"
        
        # Display each license with detailed information
        foreach ($license in $licenses) {
            $available = $license.PrepaidUnits.Enabled - $license.ConsumedUnits
            $utilizationPercent = if ($license.PrepaidUnits.Enabled -gt 0) { 
                [math]::Round(($license.ConsumedUnits / $license.PrepaidUnits.Enabled) * 100, 1) 
            } else { 0 }
            
            Write-ColorOutput "`n📋 License: $($license.SkuPartNumber)" "White"
            Write-ColorOutput "   SKU ID: $($license.SkuId)" "Gray"
            Write-ColorOutput "   Total Units: $($license.PrepaidUnits.Enabled)" "Gray"
            Write-ColorOutput "   Consumed: $($license.ConsumedUnits)" "Gray"
            Write-ColorOutput "   Available: $available" -Color $(if($available -gt 0) {"Green"} else {"Red"})
            Write-ColorOutput "   Utilization: $utilizationPercent%" "Gray"
            Write-ColorOutput "   Status: $(if ($available -gt 0) { '✅ Available' } else { '❌ No Units' })" "Gray"
        }
        
        # Create a summary table
        Write-ColorOutput "`n=== Summary Table ===" "Cyan"
        $licenseTable = @()
        
        foreach ($license in $licenses) {
            $available = $license.PrepaidUnits.Enabled - $license.ConsumedUnits
            $utilizationPercent = if ($license.PrepaidUnits.Enabled -gt 0) { 
                [math]::Round(($license.ConsumedUnits / $license.PrepaidUnits.Enabled) * 100, 1) 
            } else { 0 }
            
            $licenseInfo = [PSCustomObject]@{
                'SKU' = $license.SkuPartNumber
                'Total' = $license.PrepaidUnits.Enabled
                'Used' = $license.ConsumedUnits
                'Available' = $available
                'Utilization%' = "$utilizationPercent%"
                'Status' = if ($available -gt 0) { "✅" } else { "❌" }
            }
            
            $licenseTable += $licenseInfo
        }
        
        # Display in table format
        $licenseTable | Format-Table -AutoSize
        
        Write-ColorOutput "`n=== Common License Mappings ===" "Cyan"
        Write-ColorOutput "===============================" "Cyan"
        
        $commonMappings = @{
            "ENTERPRISEPACK" = "Office 365 E3"
            "ENTERPRISEPREMIUM" = "Office 365 E5"
            "SPE_E3" = "Microsoft 365 E3"
            "SPE_E5" = "Microsoft 365 E5"
            "DESKLESSPACK" = "Office 365 F3"
            "SPE_F1" = "Microsoft 365 F3"
            "O365_BUSINESS_PREMIUM" = "Office 365 Business Premium"
            "SPB" = "Microsoft 365 Business Premium"
        }
        
        foreach ($mapping in $commonMappings.GetEnumerator()) {
            $foundLicense = $licenses | Where-Object { $_.SkuPartNumber -eq $mapping.Key }
            if ($foundLicense) {
                $available = $foundLicense.PrepaidUnits.Enabled - $foundLicense.ConsumedUnits
                $status = if ($available -gt 0) { "✅ Available ($available units)" } else { "❌ No units available" }
                Write-Host "$($mapping.Key.PadRight(25)) → $($mapping.Value.PadRight(35)) $status" -ForegroundColor White
            }
        }
          Write-ColorOutput "`n=== Usage Instructions ===" "Yellow"
        Write-ColorOutput "To use these SKUs in the license switching script:" "White"
        Write-ColorOutput ".\Switch-Office365Licenses.ps1 -ExpiringLicenseSku ""SOURCE_SKU"" -NewLicenseSku ""TARGET_SKU"" -WhatIf" "Gray"
        Write-ColorOutput "`nExample for current tenant:" "White"
        Write-ColorOutput ".\Switch-Office365Licenses.ps1 -ExpiringLicenseSku ""ENTERPRISEPACK"" -NewLicenseSku ""SPE_E3"" -WhatIf" "Gray"
        
        if ($TenantId -or $TenantDomain) {
            Write-ColorOutput "`nFor multi-tenant scenarios:" "White"
            if ($TenantId) {
                Write-ColorOutput ".\Get-Office365Licenses.ps1 -TenantId ""$TenantId""" "Gray"
            }
            if ($TenantDomain) {
                Write-ColorOutput ".\Get-Office365Licenses.ps1 -TenantDomain ""$TenantDomain""" "Gray"
            }
        }        
        # Additional information section
        Write-ColorOutput "`n=== Additional License Information ===" "Cyan"
        Write-ColorOutput "To get more detailed information about your licenses:" "White"
        Write-ColorOutput ""
        Write-ColorOutput "1. See all service plans in each license:" "White"
        Write-ColorOutput "   Get-MgSubscribedSku | ForEach-Object { Write-Host `"License: `$(`$_.SkuPartNumber)`"; `$_.ServicePlans | Select-Object ServicePlanName, ProvisioningStatus }" "Gray"
        Write-ColorOutput ""
        Write-ColorOutput "2. See licenses assigned to a specific user:" "White"
        Write-ColorOutput "   Get-MgUserLicenseDetail -UserId 'user@domain.com'" "Gray"
        Write-ColorOutput ""
        Write-ColorOutput "3. Find all users with a specific license:" "White"
        Write-ColorOutput "   `$sku = Get-MgSubscribedSku | Where-Object {`$_.SkuPartNumber -eq 'LICENSE_SKU'}" "Gray"
        Write-ColorOutput "   Get-MgUser -Filter `"assignedLicenses/any(x:x/skuId eq `$(`$sku.SkuId))`" -ConsistencyLevel eventual" "Gray"
        
        return $licenses
    }
    catch {
        Write-ColorOutput "Failed to retrieve licenses: $($_.Exception.Message)" "Red"
        return $null
    }
}

# Main execution
Write-ColorOutput "=== Office 365 License Lookup Tool ===" "Cyan"
Write-ColorOutput "Start Time: $(Get-Date)" "Gray"

# Show connection details
if ($TenantId) {
    Write-ColorOutput "Target Tenant ID: $TenantId" "White"
}
elseif ($TenantDomain) {
    Write-ColorOutput "Target Tenant Domain: $TenantDomain" "White"
}
else {
    Write-ColorOutput "Using default tenant (interactive selection if multiple available)" "White"
}

Write-ColorOutput ""

# Connect to Microsoft Graph
if (Connect-ToMicrosoftGraph -TenantId $TenantId -TenantDomain $TenantDomain) {
    # Get and display available licenses
    $licenses = Get-AvailableLicenses
    
    if ($licenses) {
        Write-ColorOutput "`n=== License Summary ===" "Cyan"
        Write-ColorOutput "Total Subscribed License Types: $($licenses.Count)" "White"
        Write-ColorOutput "These are all the license plans your organization has purchased." "Gray"
        Write-ColorOutput "To see individual service plans within each license, use:" "Gray"
        Write-ColorOutput "Get-MgSubscribedSku | Select-Object SkuPartNumber -ExpandProperty ServicePlans" "Gray"
        
        # Show tenant context information
        $context = Get-MgContext
        if ($context) {
            Write-ColorOutput "`n=== Connected Tenant Information ===" "Cyan"
            Write-ColorOutput "Tenant ID: $($context.TenantId)" "White"
            if ($context.Account) {
                Write-ColorOutput "Connected Account: $($context.Account)" "White"
            }
        }
    }
} else {
    Write-ColorOutput "Failed to connect to Microsoft Graph. Please check your permissions and try again." "Red"
    Write-ColorOutput "`nTroubleshooting tips:" "Yellow"
    Write-ColorOutput "1. Ensure you have admin permissions in the target tenant" "Gray"
    Write-ColorOutput "2. Check if the tenant ID/domain is correct" "Gray"
    Write-ColorOutput "3. Verify your account has access to the specified tenant" "Gray"
}

Write-ColorOutput "`nEnd Time: $(Get-Date)" "Gray"
Write-ColorOutput "License lookup completed." "Green"
