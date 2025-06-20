<#
.SYNOPSIS
    Quick License Lookup Script
.DESCRIPTION
    This helper script connects to Microsoft Graph and displays all available
    Office 365 licenses with their SKUs, making it easier to identify the
    correct license names for the main switching script.
.EXAMPLE
    .\Get-Office365Licenses.ps1
#>

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
        Import-Module Microsoft.Graph.Identity.DirectoryManagement
        
        # Connect with required scopes
        Connect-MgGraph -Scopes "Directory.Read.All", "Organization.Read.All"
        
        Write-ColorOutput "Successfully connected to Microsoft Graph" "Green"
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
        
        Write-ColorOutput "`n=== Available Office 365 Licenses ===" "Cyan"
        Write-ColorOutput "=====================================" "Cyan"
        
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
        Write-ColorOutput "`nExample:" "White"
        Write-ColorOutput ".\Switch-Office365Licenses.ps1 -ExpiringLicenseSku ""ENTERPRISEPACK"" -NewLicenseSku ""SPE_E3"" -WhatIf" "Gray"
        
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
Write-ColorOutput ""

# Connect to Microsoft Graph
if (Connect-ToMicrosoftGraph) {
    # Get and display available licenses
    $licenses = Get-AvailableLicenses
    
    if ($licenses) {
        Write-ColorOutput "`nTotal licenses found: $($licenses.Count)" "Green"
    }
} else {
    Write-ColorOutput "Failed to connect to Microsoft Graph. Please check your permissions and try again." "Red"
}

Write-ColorOutput "`nEnd Time: $(Get-Date)" "Gray"
Write-ColorOutput "License lookup completed." "Green"
