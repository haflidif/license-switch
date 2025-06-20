# Office 365 License Switching Script

A comprehensive PowerShell script for bulk switching Office 365 licenses. This script connects to Microsoft Graph, validates licenses, exports users, and performs bulk license switches with proper error handling and logging.

## 🚀 Enhanced Features (Latest Version)

- ✅ **Dual Input Support**: Use either SKU Part Numbers (user-friendly) or SKU IDs (GUIDs)
- ✅ **Test Mode**: Validate functionality on limited users before full deployment
- ✅ **Optimized Performance**: Server-side filtering for fast user discovery in large tenants
- ✅ **Microsoft Graph Integration**: Automatically connects with proper scopes
- ✅ **License Validation**: Validates both source and target licenses exist and are available
- ✅ **User Export**: Exports affected users to CSV before making changes
- ✅ **Bulk Processing**: Handles multiple users efficiently with progress tracking
- ✅ **WhatIf Support**: Preview changes before execution with optional verbose output
- ✅ **Comprehensive Logging**: Colored output and detailed progress tracking
- ✅ **Error Handling**: Robust error handling with fallback mechanisms
- ✅ **Throttling Protection**: Built-in delays to prevent API throttling
- ✅ **Multi-Tenant Support**: Works across different Microsoft 365 tenants
- ✅ **Verbose Mode**: Control output verbosity with PowerShell's built-in -Verbose flag

## Prerequisites

1. **PowerShell 5.1 or later**
2. **Admin Permissions**: Global Admin or User Admin role in Microsoft 365
3. **Microsoft Graph PowerShell Module**: Will be auto-installed if missing

## Parameters

| Parameter | Required | Description | Example |
|-----------|----------|-------------|---------|
| **SkuPartNumber Mode** | | | |
| `ExpiringLicenseSku` | Yes* | SKU Part Number of license to remove | `"STANDARDPACK"` |
| `NewLicenseSku` | Yes* | SKU Part Number of license to assign | `"ENTERPRISEPACK"` |
| **SkuId (GUID) Mode** | | | |
| `ExpiringLicenseSkuId` | Yes* | SKU ID (GUID) of license to remove | `"18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e"` |
| `NewLicenseSkuId` | Yes* | SKU ID (GUID) of license to assign | `"7e31c0d9-9551-471d-836f-32ee72be4a01"` |
| **Common Parameters** | | | |
| `ExportPath` | No | Path for CSV export | `"C:\Reports\export.csv"` |
| `WhatIf` | No | Preview changes without executing | `-WhatIf` |
| `Verbose` | No | Show detailed processing output | `-Verbose` |
| **Test Mode Parameters** | | | |
| `TestMode` | No | Enable test mode for limited user processing | `-TestMode` |
| `MaxTestUsers` | No | Max users to process in test mode (default: 5) | `-MaxTestUsers 10` |
| **Multi-Tenant Parameters** | | | |
| `TenantId` | No | Azure AD Tenant ID (GUID) to connect to | `"12345678-1234-1234-1234-123456789012"` |
| `TenantDomain` | No | Azure AD Tenant domain to connect to | `"contoso.onmicrosoft.com"` |

*Note: Use either SkuPartNumber parameters OR SkuId parameters, not both.*

## Usage Examples

### 🔍 1. Discover Available Licenses (Always Start Here)
```powershell
# List all available licenses in your tenant
.\Get-Office365Licenses.ps1

# List licenses in a specific tenant
.\Get-Office365Licenses.ps1 -TenantId "your-tenant-id"
```

### 📋 2. Preview Changes (Recommended First Step)

**Using SKU Part Numbers (User-Friendly)**
```powershell
# Quiet preview (summary only)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -WhatIf

# Verbose preview (show each user)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -WhatIf -Verbose
```

**Using SKU IDs (GUIDs)**
```powershell
# Using exact GUIDs for precise targeting
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01" -WhatIf
```

**Multi-Tenant Preview**
```powershell
# Preview in specific tenant using Tenant ID
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantId "12345678-1234-1234-1234-123456789012" -WhatIf

# Preview in specific tenant using domain
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantDomain "contoso.onmicrosoft.com" -WhatIf
```

### ⚡ 3. Execute License Switch

**Standard Execution**
```powershell
# Using SKU Part Numbers
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK"

# Using SKU IDs for automation scenarios
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSkuId "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" -NewLicenseSkuId "7e31c0d9-9551-471d-836f-32ee72be4a01"
```

**Multi-Tenant Execution**
```powershell
# Execute in specific tenant using Tenant ID
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantId "12345678-1234-1234-1234-123456789012"

# Execute in specific tenant using domain
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantDomain "contoso.onmicrosoft.com"
```

**With Verbose Output**
```powershell
# Show detailed processing for each user
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -Verbose
```

### 🧪 3.5. Test Mode (HIGHLY RECOMMENDED for Large Environments)

**Why Use Test Mode?**
- Perfect for validating functionality before processing thousands of users
- Allows you to test the exact license switching process on a small subset
- Identifies potential issues early without affecting the entire user base
- Provides confidence before running on 7K+ users

**Test Mode with WhatIf (Preview Only)**
```powershell
# Test with default 5 users (safest approach)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -WhatIf

# Test with specific number of users (10 users preview)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -MaxTestUsers 10 -WhatIf

# Test Mode in specific tenant
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TenantId "12345678-1234-1234-1234-123456789012" -TestMode -WhatIf
```

**Test Mode with Actual Execution (Recommended Validation)**
```powershell
# Execute test with 3 users (recommended for initial validation)
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -MaxTestUsers 3

# Execute test with verbose output for detailed validation
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -MaxTestUsers 5 -Verbose

# Safe behavior: If you request more users than available, it uses all available users
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -TestMode -MaxTestUsers 1000 -WhatIf
```

**Test Mode Parameters**
| Parameter | Default | Description |
|-----------|---------|-------------|
| `TestMode` | Off | Enables test mode processing |
| `MaxTestUsers` | 5 | Maximum users to process in test mode. **If higher than available users, all available users are processed safely.** |

### 📁 4. Custom Export Path
```powershell
# Specify custom export location
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "STANDARDPACK" -NewLicenseSku "ENTERPRISEPACK" -ExportPath "C:\Reports\LicenseSwitch_$(Get-Date -Format 'yyyyMMdd').csv"
```

### 🏢 5. Multi-Tenant Operations
```powershell
# Work with specific tenant
.\Get-Office365Licenses.ps1 -TenantId "tenant-guid-here"
# Then use the license information for switching in that tenant
```

## 📊 Performance Enhancements

### Server-Side Filtering
The script now uses Microsoft Graph's server-side filtering for **dramatically improved performance**:

- ✅ **Before**: Downloaded ALL users, filtered locally (slow, timeouts)
- ✅ **After**: Server filters users, downloads only relevant ones (fast, reliable)

### Progress Tracking
- Real-time search duration timing
- User discovery progress indicators  
- Batch processing for large datasets
- Fallback mechanisms for reliability

## Common License Examples

### Current Microsoft 365 Licenses (2025)

| License Name | SKU Part Number | Example SKU ID |
|--------------|-----------------|----------------|
| Microsoft 365 Business Standard | `STANDARDPACK` | `f245ecc8-75af-4f8e-b61f-27d8114de5f3` |
| Office 365 Enterprise E3 | `ENTERPRISEPACK` | `6fd2c87f-b296-42f0-b197-1e91e994b900` |
| Microsoft Entra Suite | `Microsoft_Entra_Suite` | `f9602137-2203-447b-9fff-41b36e08ce5d` |
| Office 365 E3 | `ENTERPRISEPACK` | `6fd2c87f-b296-42f0-b197-1e91e994b900` |
| Office 365 E5 | `ENTERPRISEPREMIUM` | `c7df2760-2c81-4ef7-b578-5b5392b571df` |
| Microsoft 365 E3 | `SPE_E3` | `05e9a617-0261-4cee-bb44-138d3ef5d965` |
| Microsoft 365 E5 | `SPE_E5` | `06ebc4ee-1bb5-47dd-8120-11324bc54e06` |
| Microsoft 365 Business Premium | `SPB` | `cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46` |

*Note: Use `.\Get-Office365Licenses.ps1` to see current licenses available in your specific tenant.*

## 🔄 Input Methods

### Method 1: SKU Part Numbers (Recommended)
- **User-friendly names** like `"STANDARDPACK"`
- **Easier to read and understand**
- **Less prone to typos in manual operations**

### Method 2: SKU IDs (GUIDs)  
- **Exact GUIDs** like `"18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e"`
- **Perfect for automation and scripting**
- **Eliminates any ambiguity**

## Process Flow

### Enhanced Workflow (Latest Version)

1. **🔗 Connection**: Connects to Microsoft Graph with required permissions
2. **🔍 License Discovery**: Retrieves and displays all available licenses with SKU IDs
3. **✅ Validation**: Validates both source and target license SKUs (supports both formats)
4. **⚡ User Discovery**: 
   - Uses **server-side filtering** for optimal performance
   - Shows progress indicators and timing
   - Falls back to batch processing if needed
5. **📊 User Preview**: Shows sample of found users for verification
6. **📁 Export**: Exports user list to CSV file with enhanced details
7. **🤔 Confirmation**: Asks for user confirmation (unless using -WhatIf)
8. **🔄 Processing**: Switches licenses with progress tracking
9. **📈 Summary**: Provides detailed completion report

### Verbose vs Standard Output

**Standard Mode** (Default):
- Summary messages only
- Clean, focused output
- Perfect for production runs

**Verbose Mode** (`-Verbose`):
- Individual user processing details
- Detailed step-by-step progress  
- Great for troubleshooting

## Output Files

The script generates enhanced CSV export files with the following columns:
- `DisplayName`: User's display name
- `UserPrincipalName`: User's UPN/email
- `UserId`: User's unique ID
- `UsageLocation`: User's location setting
- `CurrentLicense`: The license being switched from
- `ExportDate`: Timestamp of the export

Default filename format: `LicenseSwitchExport_YYYYMMDD_HHMMSS.csv`

## Error Handling

The script includes comprehensive error handling with **enhanced fallback mechanisms**:

### Connection & Authentication
- ❌ Microsoft Graph connection failures
- ❌ Insufficient permissions
- ❌ Multi-factor authentication issues

### License Validation  
- ❌ Invalid license SKUs (both Part Numbers and IDs)
- ❌ Insufficient license availability
- ❌ License not found in tenant

### User Discovery (Enhanced)
- ❌ **Server-side filtering failures** → Automatic fallback to client-side
- ❌ **Large tenant timeouts** → Batch processing with progress tracking
- ❌ **API throttling** → Built-in retry mechanisms
- ❌ User processing errors

### Processing & Export
- ❌ User assignment failures with detailed errors
- ❌ Export file creation issues
- ❌ Partial processing scenarios

## Best Practices

### 🚀 Performance Optimization
1. **Use server-side filtering** (automatic) for large tenants
2. **Run during off-peak hours** for large user sets
3. **Monitor verbose output** during troubleshooting
4. **Test with small groups first** before bulk operations

### 🔒 Security & Compliance  
1. **Always run with -WhatIf first** to preview changes
2. **Verify license availability** before bulk operations
3. **Keep export files secure** (contain user information)
4. **Run with appropriate admin permissions only**

### 📊 Operational Excellence
1. **Use Get-Office365Licenses.ps1** to discover current licenses
2. **Keep export files** for audit purposes
3. **Use verbose mode** for detailed troubleshooting
4. **Document license changes** for compliance

## Troubleshooting

### Performance Issues

**Large Tenant Slow Searches**
```
Using optimized server-side filtering for better performance...
Server-side filtering failed: [error details]
Falling back to client-side filtering...
```
- The script automatically handles this with fallback mechanisms
- Monitor the timing messages to track performance
- Consider running during off-peak hours for very large tenants

**Server-Side Filtering Errors**
```powershell
# If you see fallback messages, the script is working as designed
# The fallback ensures compatibility across all tenant configurations
```

### License Discovery Issues

**No Licenses Found**
```
Available Licenses:
===================
(No licenses displayed)
```
- Verify admin permissions (Global Admin or User Admin required)
- Check if tenant has any purchased licenses
- Try using `-TenantId` parameter for multi-tenant scenarios

**SKU Format Issues**
```
ERROR: Invalid license SKU format
```
- Use `.\Get-Office365Licenses.ps1` to see exact SKU formats
- Choose either SKU Part Number OR SKU ID format (don't mix)
- Copy exact names from the license discovery output

### Getting Help

1. **Start with license discovery**: `.\Get-Office365Licenses.ps1`
2. **Use WhatIf for previews**: Add `-WhatIf` to any command
3. **Enable verbose output**: Add `-Verbose` for detailed processing
4. **Check the CSV export** for affected users
5. **Review colored output** for specific error messages

## 🛠️ Available Scripts

| Script | Purpose | Key Features |
|--------|---------|--------------|
| `Switch-Office365Licenses.ps1` | **Main license switching script** | Dual input, performance optimized, verbose mode |
| `Get-Office365Licenses.ps1` | **License discovery and inspection** | Multi-tenant, service plan details, extra commands |
| `runbook.ps1` | **Azure Automation integration** | Automated license switching in Azure |
| `Examples.ps1` | **Usage examples and scenarios** | Real-world automation examples |

## 🔧 Integration Options

### Azure Automation
Use `runbook.ps1` for automated license switching:
- Scheduled license transitions
- Automated compliance workflows  
- Integration with other Azure services

### Multi-Tenant Management
Use tenant-specific parameters:
```powershell
.\Get-Office365Licenses.ps1 -TenantId "your-tenant-guid"
```

### CI/CD Integration
Perfect for DevOps scenarios:
- Use SKU IDs for exact targeting
- WhatIf mode for validation
- CSV exports for audit trails

## Security Notes

- **Elevated permissions required**: Global Admin or User Admin
- **All operations logged**: Comprehensive audit trail
- **User consent required**: Confirmation before changes (unless WhatIf)
- **Secure file handling**: Export files contain user information
- **MFA compatible**: Works with modern authentication

## Version History

### Latest Enhancements (June 2025)
- ✅ **Dual input support**: SKU Part Numbers and SKU IDs
- ✅ **Performance optimization**: Server-side filtering
- ✅ **Verbose mode control**: Standard vs detailed output
- ✅ **Enhanced error handling**: Fallback mechanisms
- ✅ **Improved user experience**: Progress tracking and timing

## Support

For issues or questions:
1. **Check troubleshooting section** above
2. **Review PowerShell execution policy** requirements  
3. **Verify Microsoft Graph module** installation
4. **Ensure proper Office 365 admin permissions**
5. **Test with Get-Office365Licenses.ps1** first
6. **Use -WhatIf and -Verbose** for diagnostics
