# Get-SPOFileswithLabels.ps1

## Overview

**Get-SPOFileswithLabels.ps1** is a PowerShell script that searches for files (such as PDF, DOCX, XLSX) in SharePoint Online and OneDrive for Business using the Microsoft Graph API. It retrieves file metadata, sensitivity labels, and retention labels, then exports the results to a CSV file. The script is robust, supporting certificate and client secret authentication, advanced throttling and pagination handling, and detailed logging.

## Features

- Authenticates with Microsoft Graph (client secret or certificate)
- Searches OneDrive and SharePoint Online for files of a specified type
- Extracts and exports:
  - File details (name, URL, size, created/modified dates, etc.)
  - Sensitivity labels
  - Retention labels
- Handles large result sets with batching and pagination
- Implements exponential backoff for throttling
- Supports verbose debug output

## Prerequisites

- PowerShell 5.1 or higher
- Registered Azure AD application with appropriate Microsoft Graph API permissions:
  - Sites.Read.All
  - Files.Read.All
  - InformationProtectionPolicy.Read.All
  - RecordsManagement.Read.All
- Either a client secret or a certificate for authentication to Microsoft Graph

## Permissions Required

- Sites.Read.All
- Files.Read.All
- InformationProtectionPolicy.Read.All
- RecordsManagement.Read.All

## Configuration

Update the variables at the top of the script to match your environment:

- `$tenantName` — Your SharePoint tenant (e.g., "contoso")
- `$tenantId` — Azure AD tenant ID
- `$clientId` — Application (client) ID from Azure AD
- `$clientSecret` — Application client secret (if using client secret auth)
- `$Thumbprint` — Certificate thumbprint (if using certificate auth)
- `$CertStore` — Certificate store: 'LocalMachine' or 'CurrentUser'
- `$fileType` — File extension to search for (e.g., 'pdf', 'xlsx')
- `$debug` — Set to `$true` for detailed output
- `$batchSize`, `$delayBetweenBatches`, `$maxConcurrentRequests` — Tune API batching and rate limiting as needed

## Usage

1. Edit the configuration section at the top of the script with your tenant and authentication details.
2. Ensure your Azure AD app registration has the required Microsoft Graph permissions and admin consent.
3. Open a PowerShell session and run:
   ```powershell
   .\Get-SPOFileswithLabels.ps1
   ```
4. The script will authenticate, search for the specified file type, and export results to a CSV file in your `%TEMP%` directory.

## Output

- CSV file named `SPOFileswithLabels_Search_Results_<timestamp>.csv` in `%TEMP%` directory.
- Columns include:
  - File name
  - File URL
  - Created/modified dates
  - Size
  - Sensitivity label
  - Retention label
  - Other metadata as available

## Advanced Info

- Handles throttling and server errors using exponential backoff and retry logic
- Preloads and caches all sensitivity labels for efficiency
- Supports both OneDrive personal/user and SharePoint teams/sites paths
- Uses Microsoft Graph endpoints for search and label extraction

## References

- [Microsoft Graph Search API Overview](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview)
- [DriveItem: extractSensitivityLabels](https://learn.microsoft.com/en-us/graph/api/driveitem-extractsensitivitylabels?view=graph-rest-1.0&tabs=http)

## Troubleshooting

- Ensure your Azure AD app registration has all required permissions and admin consent
- For certificate authentication, verify the certificate is available in the specified store
- For throttling errors, the script will retry; adjust batch size and delay if needed
- Set `$debug = $true` for more detailed logging
- Review error messages in the PowerShell output for further troubleshooting

## Author

- Mike Lee

---

**Disclaimer:** This script is provided as-is without warranty. Use at your own risk. Contributions and feedback are welcome!
