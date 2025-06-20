# Office 365 License Switch Examples - Enhanced Version
# Copy and modify these examples for your specific needs
# Version: Enhanced (June 2025) - Supports dual input methods and performance optimization

# ================================================
# STEP 1: License Discovery (ALWAYS START HERE!)
# ================================================

# Discover all available licenses in your current tenant
.\Get-Office365Licenses.ps1

# Discover licenses in a specific tenant (multi-tenant scenarios)
.\Get-Office365Licenses.ps1 -TenantId "49ff7219-653a-4644-8540-71d16dbf9c16"
.\Get-Office365Licenses.ps1 -TenantDomain "contoso.onmicrosoft.com"

# ================================================
# STEP 2: Preview License Switches (CRITICAL!)
# ================================================

# === Using SKU Part Numbers (User-Friendly Method) ===

# Example 1: Microsoft 365 E5 (no Teams) to Teams Enterprise
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -WhatIf

# Example 2: Office 365 E3 to Microsoft 365 E3
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3" -WhatIf

# Example 3: Office 365 E5 to Microsoft 365 E5  
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPREMIUM" -NewLicenseSku "SPE_E5" -WhatIf

# === Using SKU IDs (GUID Method - Perfect for Scripting) ===

# Example 1: Same switch using exact GUIDs
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01" -WhatIf

# Example 2: E3 to E5 upgrade using GUIDs
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "6fd2c87f-b296-42f0-b197-1e91e994b900" -NewLicenseSkuId "c7df2760-2c81-4ef7-b578-5b5392b571df" -WhatIf

# === Verbose Preview (Show Each User) ===

# Detailed preview showing individual user processing
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -WhatIf -Verbose

# === Multi-Tenant Preview Examples ===

# Preview in specific tenant using Tenant ID
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantId "12345678-1234-1234-1234-123456789012" -WhatIf

# Preview in specific tenant using domain  
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantDomain "contoso.onmicrosoft.com" -WhatIf

# Preview with SKU IDs in specific tenant
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01" -TenantId "12345678-1234-1234-1234-123456789012" -WhatIf

# ================================================
# STEP 2.5: Test Mode (NEW!) - HIGHLY RECOMMENDED  
# ================================================

# === Test Mode for Large Environments ===
# Perfect for validating functionality before processing thousands of users!

# Test with default 5 users (safest approach)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -WhatIf

# Test with specific number of users (10 users preview)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -MaxTestUsers 10 -WhatIf

# Test Mode with SKU IDs (scripting testing)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01" -TestMode -MaxTestUsers 3 -WhatIf

# Test Mode in specific tenant
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantId "12345678-1234-1234-1234-123456789012" -TestMode -WhatIf

# === Actual Test Execution (Limited Users) ===
# Execute the license switch on a small subset to validate functionality

# Execute test with 3 users (recommended for initial validation)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -MaxTestUsers 3

# Execute test with verbose output for detailed validation
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -MaxTestUsers 5 -Verbose

# Execute test in specific tenant with 2 users
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantDomain "contoso.onmicrosoft.com" -TestMode -MaxTestUsers 2

# === Safe Behavior: Requested Users > Available Users ===
# When you specify more test users than actually exist, it gracefully uses all available users

# If only 100 users exist but you request 120, it will test all 100 users (safe behavior)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -MaxTestUsers 120 -WhatIf

# Large number for "test all available users" approach
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -MaxTestUsers 1000 -WhatIf

# ================================================
# STEP 3: Execute License Switches  
# ================================================

# === Standard Execution (Clean Output) ===

# Example 1: Execute M365 E5 (no Teams) to Teams Enterprise
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK"

# Example 2: Execute using SKU IDs (great for automated scripts)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01"

# === Multi-Tenant Execution ===

# Execute in specific tenant using Tenant ID
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantId "12345678-1234-1234-1234-123456789012"

# Execute in specific tenant using domain
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantDomain "contoso.onmicrosoft.com"

# Execute using SKU IDs in specific tenant
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01" -TenantId "12345678-1234-1234-1234-123456789012"

# === Verbose Execution (Detailed Output) ===

# Execute with detailed processing information
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -Verbose

# === Custom Export Paths ===

# Execute with custom export location
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -ExportPath "C:\Reports\LicenseSwitch_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

# ================================================
# ENHANCED LICENSE MAPPINGS (2025)
# ================================================

<#
=== Current Microsoft 365 License Examples ===

License Name                    | SKU Part Number                | Example SKU ID (GUID)
------------------------------- | ------------------------------ | --------------------------------------
Microsoft 365 Business Standard| STANDARDPACK                   | f245ecc8-75af-4f8e-b61f-27d8114de5f3
Office 365 Enterprise E3       | ENTERPRISEPACK                 | 6fd2c87f-b296-42f0-b197-1e91e994b900
Microsoft Entra Suite          | Microsoft_Entra_Suite          | f9602137-2203-447b-9fff-41b36e08ce5d
Office 365 E3                  | ENTERPRISEPACK                 | 6fd2c87f-b296-42f0-b197-1e91e994b900
Office 365 E5                  | ENTERPRISEPREMIUM              | c7df2760-2c81-4ef7-b578-5b5392b571df
Microsoft 365 E3               | SPE_E3                         | 05e9a617-0261-4cee-bb44-138d3ef5d965
Microsoft 365 E5               | SPE_E5                         | 06ebc4ee-1bb5-47dd-8120-11324bc54e06
Microsoft 365 Business Premium | SPB                            | cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46

=== Common Migration Scenarios ===

Scenario                        | From                          | To
------------------------------- | ----------------------------- | -------------------------------
Standard to Enterprise         | STANDARDPACK                  | ENTERPRISEPACK
Office to Microsoft 365       | ENTERPRISEPACK                | SPE_E3
License Upgrades               | SPE_E3                        | SPE_E5  
Business to Enterprise         | SPB                           | SPE_E3
Legacy to Modern              | ENTERPRISEPREMIUM              | SPE_E5

#>

# ================================================
# PERFORMANCE TESTING & OPTIMIZATION
# ================================================

# Test server-side filtering performance (automatic in latest version)
# The script will show timing information like:
# "User search completed in 0.67 seconds"
# "Using optimized server-side filtering for better performance..."

# For large tenants, monitor the progress:
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -WhatIf -Verbose

# ================================================
# TROUBLESHOOTING & DIAGNOSTICS
# ================================================

# === License Discovery Issues ===

# Check specific tenant licenses
.\Get-Office365Licenses.ps1 -TenantId "your-tenant-guid-here"

# === Performance Issues ===

# If you see "Server-side filtering failed" messages:
# The script automatically falls back to client-side filtering
# This is normal and ensures compatibility

# === User Discovery Issues ===

# Check if users actually have the license:
# Get-MgUser -UserId "user@domain.com" -Property AssignedLicenses | Select-Object AssignedLicenses

# Check available license units:
# Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq "Microsoft_Teams_Enterprise_New"} | Select-Object SkuPartNumber, @{Name="Available"; Expression={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}}

# ================================================
# BATCH PROCESSING EXAMPLES
# ================================================

# Sequential processing for multiple license types:

# Phase 1: Preview all changes first
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -WhatIf
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3" -WhatIf

# Phase 2: Execute changes after verification
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK"
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3"

# ================================================
# MULTI-TENANT SCENARIOS
# ================================================

# Working with multiple tenants:
$tenants = @(
    "tenant1-guid-here",
    "tenant2-guid-here"
)

foreach ($tenant in $tenants) {
    Write-Host "Processing tenant: $tenant" -ForegroundColor Yellow
    .\Get-Office365Licenses.ps1 -TenantId $tenant
    # Review licenses, then execute switches as needed
}

# ================================================
# SAFETY & COMPLIANCE CHECKLIST
# ================================================

<#
✅ PRE-EXECUTION CHECKLIST:
1. Run Get-Office365Licenses.ps1 to discover available licenses
2. ALWAYS use -WhatIf first to preview changes
3. Verify sufficient target license availability  
4. Test with a small user group first
5. Plan execution during maintenance windows
6. Ensure proper admin permissions (Global Admin or User Admin)

✅ DURING EXECUTION:
1. Monitor the console output for errors
2. Watch for performance messages and timing
3. Note any fallback mechanism activations
4. Verify CSV export file creation

✅ POST-EXECUTION:
1. Review the summary report
2. Keep CSV export files for audit purposes
3. Verify users have correct licenses applied
4. Document the changes for compliance
5. Monitor user experience for any issues

✅ OUTPUT MODES:
- Standard: Clean summary output for production
- Verbose (-Verbose): Detailed processing for troubleshooting
- WhatIf (-WhatIf): Preview mode without making changes
#>

# ================================================
# ADVANCED SCRIPTING EXAMPLES
# ================================================

# Example 1: Manual license switching with error handling
try {
    # Preview first
    $whatIfResult = .\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01" -WhatIf
    
    if ($whatIfResult) {
        # If preview looks good, execute
        .\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01"
    }
} catch {
    Write-Error "License switch failed: $($_.Exception.Message)"
}

# Example 2: Conditional execution based on license availability
$licenses = .\Get-Office365Licenses.ps1
$teamsLicense = $licenses | Where-Object {$_.SkuPartNumber -eq "Microsoft_Teams_Enterprise_New"}

if ($teamsLicense.Available -gt 0) {
    Write-Host "Teams licenses available: $($teamsLicense.Available)" -ForegroundColor Green
    .\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -WhatIf
} else {
    Write-Host "No licenses available. Purchase more licenses first." -ForegroundColor Red
}
