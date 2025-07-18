# SharePoint & OneDrive Microsoft Graph Search Scripts

This repository contains PowerShell scripts for querying and reporting on SharePoint Online and OneDrive data using the Microsoft Graph API. These scripts are designed for administrative and security/compliance reporting across Microsoft 365 tenants.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Authentication](#authentication)
- [Script Descriptions](#script-descriptions)
  - [TenantWide_SP_Copilot_Agents.ps1](#tenantwide_sp_copilot_agentsps1)
  - [TenantWide_SP_Copilot_Agents_Insights.ps1](#tenantwide_sp_copilot_agents_insightsps1)
  - [search-spo_odb.ps1](#search-spo_odbps1)
  - [search_and_download.ps1](#search_and_downloadps1)
  - [Get-SPOFileswithLabels.ps1](#get-spofileswithlabelsps1)
  - [Convert-SPOfiles.ps1](#convert-spofilesps1)
- [Disclaimer](#disclaimer)

---

## Overview

These scripts use the Microsoft Graph API to:
- Search for SharePoint Copilot Agents across your entire Microsoft 365 tenant.
- Generate detailed CSV and log reports about Copilot Agents, including security and site metadata.
- Perform bulk searches on SharePoint/OneDrive items using custom queries.
- Download files found via search queries, if desired.
- Extract sensitivity labels from OneDrive/SharePoint files.
- Convert legacy Office files to modern formats (docx, xlsx, pptx, pdf) across SharePoint and OneDrive.

All scripts support secure authentication and pagination of large result sets.

---

## Prerequisites

- **PowerShell 7.x** recommended (PowerShell 5.1 or higher required for some scripts).
- **PnP.PowerShell module** (for insights script):  
  Install with `Install-Module -Name PnP.PowerShell -Scope CurrentUser`
- A **registered Azure AD/Entra application** with permissions:
  - `Sites.Read.All`
  - `Sites.FullControl.All` (for insights)
  - `Files.Read.All` (for OneDrive/SharePoint search)
  - `Files.ReadWrite.All` (for file conversion)
  - `SearchQuery.All`
  - `InformationProtectionPolicy.Read.All` (for sensitivity labels)
- **Certificate-based authentication** recommended for the Insights script and supported in conversion script.
- Appropriate admin rights for the SharePoint tenant and the Entra app registration.

---

## Authentication

All scripts use OAuth 2.0 authentication via Azure AD application credentials.  
You must set these variables at the top of each script:

```powershell
$tenantId = '<your-tenant-id>'
$clientId = '<your-client-id>'
$clientSecret = '<your-client-secret>' # Not needed if using certificate auth
```

For the _Insights_ and _Convert-SPOfiles_ scripts, certificate-based authentication is supported.  
Set these variables at the top:

```powershell
$tenantname = "<your-tenant-name>"  # Without .onmicrosoft.com
$appID = "<your-app-id>"
$thumbprint = "<your-cert-thumbprint>"
$tenantid = "<your-tenant-id>"
```

---

## Script Descriptions

### TenantWide_SP_Copilot_Agents.ps1

**Purpose:**  
Searches all SharePoint and OneDrive locations in the tenant for files of type 'agent' (SharePoint Copilot Agents) using Microsoft Graph Search API. Exports summary results to a timestamped CSV file.

**Parameters:**
- `tenantId`, `clientId`, `clientSecret`: Azure AD app credentials
- `searchRegion` (default: `"NAM"`): Region scope for search

**Output:**  
- CSV file in your `%TEMP%` directory, named `TenantWide_SharePoint_Agents<timestamp>.csv`.  
  Columns: `ID`, `Name`, `WebURL`, `CreatedDate`, `LastAccessedDate`, `Owner`

**Usage Example:**
```powershell
# Set credentials and run
./TenantWide_SP_Copilot_Agents.ps1
```

---

### TenantWide_SP_Copilot_Agents_Insights.ps1

**Purpose:**  
Discovers and documents all SharePoint Copilot Agents across the tenant, producing a **detailed report** with agent and host site metadata, security settings, sensitivity labels, and site owner info.

**Parameters:**
- `tenantname`: Your tenant’s short name (no domain)
- `appID`: Entra (Azure AD) App ID
- `thumbprint`: Certificate thumbprint for authentication
- `tenantid`: Tenant GUID
- `searchRegion`: `"NAM"`, `"EMEA"`, etc.

**Features:**
- Uses Microsoft Graph for agent search
- Uses PnP.PowerShell for detailed site reporting
- Handles throttling and logs all actions
- Outputs both CSV and log files

**Output:**
- CSV file: `SPO_Copilot_Agents_<timestamp>.csv` in `%TEMP%`
- Log file: `SPO_Copilot_Agents_<timestamp>.log` in `%TEMP%`

**Usage Example:**
```powershell
# Set variables at the top, then run:
.\TenantWide_SP_Copilot_Agents_Insights.ps1
```

---

### search-spo_odb.ps1

**Purpose:**  
Performs custom search queries on OneDrive (and SharePoint) items via the Graph API. Reads each search query from an external file and exports results to CSV.

**Parameters:**
- `tenantId`, `clientId`, `clientSecret`: Azure AD app credentials
- `searchRegion`: Search region (default `"NAM"`)
- `searchQueryList`: Path to a plain text file with one query per line (e.g., `C:\temp\userlist.txt`)
- `LogName`: Path to the output CSV (default: auto-generated in `%TEMP%`)

**Output:**
- CSV file in `%TEMP%` named `Search_Results_<timestamp>.csv`

**Usage Example:**
```powershell
# Prepare a plain text file with search queries, one per line
$searchQueryList = Get-Content 'C:\temp\userlist.txt'
# Run the script
./search-spo_odb.ps1
```

---

### search_and_download.ps1

**Purpose:**  
Performs custom search queries on OneDrive items, **downloads matching files**, and exports results to CSV.

**Parameters:**
- `tenantId`, `clientId`, `clientSecret`: Azure AD app credentials
- `searchRegion`: Search region (default `"NAM"`)
- `searchQueryList`: Path to a file with queries (e.g., `C:\temp\userlist.txt`)
- `LogName`: Output CSV path (default: auto-generated in `%TEMP%`)
- `downloadPath`: Local folder to store downloaded files (default: `C:\temp\`)

**Output:**
- CSV file in `%TEMP%` named `Search_Results_<timestamp>.csv`
- Downloaded files in specified `downloadPath`

**Usage Example:**
```powershell
# Prepare a search query file
$searchQueryList = Get-Content 'C:\temp\userlist.txt'
# Run the script
./search_and_download.ps1
```

---

### Get-SPOFileswithLabels.ps1

**Purpose:**  
This Script searches OneDrive or SharePoint for files of a specified type (e.g., PDF, DOCX) and extracts details and sensitivity label information using Microsoft Graph API. Exports results to a CSV with file metadata and label info.

**Features:**
- Authenticates with Microsoft Graph via client secret or certificate
- Supports file type filtering (e.g., pdf, docx, xlsx)
- Handles throttling and pagination for large result sets
- Extracts sensitivity labels and key file metadata
- Outputs results to a timestamped CSV file

**Parameters:**
- `tenantName`: Tenant short name (e.g., "contoso")
- `tenantId`, `clientId`, `clientSecret`: Azure AD app credentials
- `AuthType`: `'ClientSecret'` or `'Certificate'`
- `Thumbprint`, `CertStore`: Certificate details (if using certificate auth)
- `fileType`: File extension to search for (e.g., `"pdf"`)
- `searchRegion`: Search region (default `"NAM"`)
- `debug`: Verbose output (`$true`/`$false`)

**Output:**  
- CSV file in your `%TEMP%` directory, named `SPOFileswithLabels_Search_Results_<timestamp>.csv`.  
  Columns:  
  - `ID`
  - `Name`
  - `WebURL`
  - `CreatedDate`
  - `LastAccessedDate`
  - `Owner`
  - `SensitivityLabel`

**Usage Example:**
```powershell
# Set your credentials and desired fileType at the top of the script
./Get-SPOFileswithLabels.ps1
```

**Prerequisites:**
- PowerShell 5.1 or higher
- Azure AD application with API permissions:
  - `Sites.Read.All`
  - `Files.Read.All`
  - `InformationProtectionPolicy.Read.All`
- Microsoft Graph API access

**Documentation:**  
- [Microsoft Graph Search API Overview](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview)
- [Extract Sensitivity Labels API](https://learn.microsoft.com/en-us/graph/api/driveitem-extractsensitivitylabels?view=graph-rest-1.0&tabs=http)

---

### Convert-SPOfiles.ps1

**Purpose:**  
Searches tenant-wide SharePoint and OneDrive locations for Office files and converts them to modern Office formats (docx, xlsx, pptx, pdf) using Microsoft Graph API. Converts supported file types and (optionally) uploads converted files back to SharePoint. Results are exported to a CSV file with metadata and conversion status.

**Parameters:**
- `tenantId`, `clientId`, `clientSecret`: Azure AD app credentials (or certificate for auth)
- `AuthType`: `'ClientSecret'` or `'Certificate'`
- `Thumbprint`, `CertStore`: Certificate details (if using certificate auth)
- `searchRegion`: Region scope for search (default `"NAM"`)
- `searchUrl`: SharePoint path to search within (default provided, can be changed)
- `fileType`: File extension(s) to search for (e.g., `"docx"`, `"xlsx"`, or multiple: `"doc,docx"`)
- `convertTo`: Target format for conversion (`"docx"`, `"xlsx"`, `"pptx"`, `"pdf"`)
- `downloadFolder`: Local path to store converted files (default: temp directory)

**Features:**
- Authenticates with Microsoft Graph using client secret or certificate
- Searches for convertible Office files across SharePoint/OneDrive
- Converts files to modern formats using Graph API or local Excel COM (for CSV)
- Handles pagination and includes robust error handling
- Uploads converted files back to SharePoint (same folder)
- Exports results with file metadata, conversion status, and upload URL

**Output:**  
- CSV file in your `%TEMP%` directory, named `TenantWide_SharePoint_FileConversion_<timestamp>.csv`
- Converted files in a timestamped subfolder in the specified download directory
- (Optional) Converted files uploaded back to SharePoint

**Usage Example:**
```powershell
# Set your credentials and desired parameters at the top of the script
./Convert-SPOfiles.ps1
```

**Prerequisites:**
- PowerShell 5.1 or higher
- Azure AD application with API permissions:
  - `Files.Read.All` or `Files.ReadWrite.All`
  - `Sites.Read.All` or `Sites.ReadWrite.All`
- (Optional) Microsoft Excel installed for local CSV conversion

**Documentation:**  
- [Microsoft Graph File Conversion API](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format)
- [Microsoft Graph Search API Overview](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview)

---

## Disclaimer

The sample scripts are provided **AS IS** without warranty of any kind.  
Microsoft and the script authors disclaim all implied warranties including, without limitation, warranties of merchantability or fitness for a particular purpose.  
The entire risk arising out of the use or performance of the scripts and documentation remains with you.  
In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever.

---

