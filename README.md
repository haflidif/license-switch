# Office 365 License Switching Script

A comprehensive PowerShell script for bulk switching Office 365 licenses. This script connects to Microsoft Graph, validates licenses, exports users, and performs bulk license switches with proper error handling and logging.

## Features

- ✅ **Microsoft Graph Integration**: Automatically connects with proper scopes
- ✅ **License Validation**: Validates both source and target licenses exist and are available
- ✅ **User Export**: Exports affected users to CSV before making changes
- ✅ **Bulk Processing**: Handles multiple users efficiently
- ✅ **WhatIf Support**: Preview changes before execution
- ✅ **Comprehensive Logging**: Colored output and detailed progress tracking
- ✅ **Error Handling**: Robust error handling with detailed error messages
- ✅ **Throttling Protection**: Built-in delays to prevent API throttling

## Prerequisites

1. **PowerShell 5.1 or later**
2. **Admin Permissions**: Global Admin or User Admin role in Microsoft 365
3. **Microsoft Graph PowerShell Module**: Will be auto-installed if missing

## Parameters

| Parameter | Required | Description | Example |
|-----------|----------|-------------|---------|
| `ExpiringLicenseSku` | Yes | SKU of the license to remove | `"ENTERPRISEPACK"` |
| `NewLicenseSku` | Yes | SKU of the license to assign | `"SPE_E3"` |
| `ExportPath` | No | Path for CSV export | `"C:\Reports\export.csv"` |
| `WhatIf` | No | Preview changes without executing | `-WhatIf` |

## Usage Examples

### 1. Preview Changes (Recommended First Step)
```powershell
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3" -WhatIf
```

### 2. Execute License Switch
```powershell
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3"
```

### 3. Custom Export Path
```powershell
.\Switch-Office365Licenses.ps1 -ExpiringLicenseSku "ENTERPRISEPACK" -NewLicenseSku "SPE_E3" -ExportPath "C:\Reports\LicenseSwitch_$(Get-Date -Format 'yyyyMMdd').csv"
```

## Common License SKUs

| License Name | SKU |
|--------------|-----|
| Office 365 E3 | `ENTERPRISEPACK` |
| Office 365 E5 | `ENTERPRISEPREMIUM` |
| Microsoft 365 E3 | `SPE_E3` |
| Microsoft 365 E5 | `SPE_E5` |
| Office 365 F3 | `DESKLESSPACK` |
| Microsoft 365 F3 | `SPE_F1` |
| Office 365 Business Premium | `O365_BUSINESS_PREMIUM` |
| Microsoft 365 Business Premium | `SPB` |

## Process Flow

1. **Connection**: Connects to Microsoft Graph with required permissions
2. **License Discovery**: Retrieves and displays all available licenses
3. **Validation**: Validates both source and target license SKUs
4. **User Discovery**: Finds all users with the expiring license
5. **Export**: Exports user list to CSV file
6. **Confirmation**: Asks for user confirmation (unless using -WhatIf)
7. **Processing**: Switches licenses for each user
8. **Summary**: Provides detailed completion report

## Output Files

The script generates a CSV export file with the following columns:
- `DisplayName`: User's display name
- `UserPrincipalName`: User's UPN/email
- `UserId`: User's unique ID
- `CurrentLicense`: The license being switched from
- `ExportDate`: Timestamp of the export

Default filename format: `LicenseSwitchExport_YYYYMMDD_HHMMSS.csv`

## Error Handling

The script includes comprehensive error handling for:
- ❌ Microsoft Graph connection failures
- ❌ Invalid license SKUs
- ❌ Insufficient license availability
- ❌ User processing errors
- ❌ Export failures

## Best Practices

1. **Always run with -WhatIf first** to preview changes
2. **Verify license availability** before bulk operations
3. **Run during maintenance windows** for large user sets
4. **Keep export files** for audit purposes
5. **Monitor the output** for any failures during processing

## Troubleshooting

### Common Issues

**Connection Errors**
```
Failed to connect to Microsoft Graph: Access denied
```
- Ensure you have Global Admin or User Admin permissions
- Check if your account has MFA properly configured

**License Not Found**
```
ERROR: New license SKU 'INVALID_SKU' not found!
```
- Run the script first to see available licenses
- Use the exact SKU name (case-sensitive)

**No Available Units**
```
WARNING: New license 'SPE_E5' has no available units!
```
- Purchase more licenses or free up existing ones
- Check license consumption in the admin center

### Getting Help

1. Run the script without parameters to see available licenses
2. Use `-WhatIf` to preview all changes
3. Check the CSV export for affected users
4. Review the colored output for specific error messages

## Security Notes

- The script requires elevated permissions
- All operations are logged and tracked
- User consent is required before making changes
- Export files contain user information - handle securely

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the PowerShell execution policy requirements
3. Verify Microsoft Graph module installation
4. Ensure proper Office 365 admin permissions
