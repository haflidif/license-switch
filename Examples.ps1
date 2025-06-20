# Office 365 License Switch Examples
# Copy and modify these examples for your specific needs

# ===========================================
# STEP 1: First, discover available licenses
# ===========================================
# Run this to see all available licenses in your tenant
.\Get-Office365Licenses.ps1

# ===========================================
# STEP 2: Preview the license switch (ALWAYS DO THIS FIRST!)
# ===========================================

# Example 1: Switch from Office 365 E3 to Microsoft 365 E3
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3" -WhatIf

# Example 2: Switch from Office 365 E5 to Microsoft 365 E5
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPREMIUM" -NewLicenseSku "SPE_E5" -WhatIf

# Example 3: Switch from Office 365 Business Premium to Microsoft 365 Business Premium
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "O365_BUSINESS_PREMIUM" -NewLicenseSku "SPB" -WhatIf

# ===========================================
# STEP 3: Execute the actual license switch
# ===========================================

# Example 1: Execute Office 365 E3 to Microsoft 365 E3 switch
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3"

# Example 2: Execute with custom export path
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3" -ExportPath "C:\Reports\E3_to_M365E3_Switch.csv"

# ===========================================
# COMMON LICENSE MAPPINGS
# ===========================================
<#
Office 365 → Microsoft 365 Migrations:

Old License (Office 365)     → New License (Microsoft 365)
ENTERPRISEPACK              → SPE_E3              (E3 to E3)
ENTERPRISEPREMIUM           → SPE_E5              (E5 to E5)
O365_BUSINESS_PREMIUM       → SPB                 (Business Premium)
DESKLESSPACK                → SPE_F1              (F3 to F3)

License Upgrades:
ENTERPRISEPACK              → SPE_E5              (E3 to E5)
SPE_E3                      → SPE_E5              (E3 to E5)
O365_BUSINESS_PREMIUM       → SPE_E3              (Business to E3)

License Downgrades:
SPE_E5                      → SPE_E3              (E5 to E3)
SPE_E3                      → SPB                 (E3 to Business)
#>

# ===========================================
# TROUBLESHOOTING COMMANDS
# ===========================================

# If you need to check specific user licenses before switching:
# Get-MgUser -UserId "user@domain.com" -Property AssignedLicenses | Select-Object AssignedLicenses

# If you need to manually check license availability:
# Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq "SPE_E3"} | Select-Object SkuPartNumber, @{Name="Available"; Expression={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}}

# ===========================================
# BATCH PROCESSING EXAMPLES
# ===========================================

# For multiple license switches, run them sequentially:
# Switch from O365 E3 to M365 E3
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3"

# Then switch from O365 E5 to M365 E5
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPREMIUM" -NewLicenseSku "SPE_E5"

# ===========================================
# SAFETY REMINDERS
# ===========================================
<#
1. ALWAYS run with -WhatIf first!
2. Verify you have enough target licenses available
3. Run during maintenance windows for large batches
4. Keep the CSV export files for audit purposes
5. Test with a small group first if possible
6. Ensure you have proper admin permissions
#>
