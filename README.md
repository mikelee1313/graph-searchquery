Certainly! Here is a detailed README for the scripts you provided in your repo **mikelee1313/graph-searchquery**. This README explains the purpose, usage, parameters, prerequisites, and output of each script, organized for clarity and ease of use.

---

# SharePoint & OneDrive Microsoft Graph Search Scripts

This repository contains PowerShell scripts for querying and reporting on SharePoint Online and OneDrive data using the Microsoft Graph API. These scripts are designed for administrative and security/compliance teams to search, inventory, and download data related to SharePoint Copilot Agents and general OneDrive/SharePoint content.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Authentication](#authentication)
- [Script Descriptions](#script-descriptions)
  - [TenantWide_SP_Copilot_Agents.ps1](#tenantwide_sp_copilot_agentsps1)
  - [TenantWide_SP_Copilot_Agents_Insights.ps1](#tenantwide_sp_copilot_agents_insightsps1)
  - [search-spo_odb.ps1](#search-spo_odbps1)
  - [search_and_download.ps1](#search_and_downloadps1)
- [Disclaimer](#disclaimer)
- [Authors](#authors)

---

## Overview

These scripts use the Microsoft Graph API to:
- Search for SharePoint Copilot Agents across your entire Microsoft 365 tenant.
- Generate detailed CSV and log reports about Copilot Agents, including security and site metadata.
- Perform bulk searches on SharePoint/OneDrive items using custom queries.
- Download files found via search queries, if desired.

All scripts support secure authentication and pagination of large result sets.

---

## Prerequisites

- **PowerShell 7.x** recommended.
- **PnP.PowerShell module** (for insights script):  
  Install with `Install-Module -Name PnP.PowerShell -Scope CurrentUser`
- A **registered Azure AD/Entra application** with permissions:
  - `Sites.Read.All`
  - `Sites.FullControl.All` (for insights)
  - `Files.Read.All` (for OneDrive/SharePoint search)
  - `SearchQuery.All`
- **Certificate-based authentication** recommended for the Insights script.
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

For the _Insights_ script, certificate-based authentication is used.  
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

## Disclaimer

The sample scripts are provided **AS IS** without warranty of any kind.  
Microsoft and the script authors disclaim all implied warranties including, without limitation, warranties of merchantability or fitness for a particular purpose.  
The entire risk arising out of the use or performance of the scripts and documentation remains with you.  
In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever.

---

## Authors

- Mike Lee
- Chanchal Jain

For questions, open an issue or contact the authors via GitHub.

---

**Happy searching and reporting!**

---

Let me know if you want a shorter README, or if you need examples or troubleshooting sections added.
