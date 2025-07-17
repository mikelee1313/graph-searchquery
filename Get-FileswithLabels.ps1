<#
.SYNOPSIS
Searches for PDF files in OneDrive and extracts their sensitivity labels using Microsoft Graph API.

.DESCRIPTION
This script authenticates with Microsoft Graph API, performs a search for specific files in OneDrive,
and extracts detailed information including sensitivity labels. The results are exported to a CSV file.

.PARAMETER None
This script does not accept parameters through the command line. Configuration is done through variables
at the beginning of the script.

.NOTES
File Name       : Get-FileswithLabels.ps1
Author          : Mike Lee
Date Created    : 7/17/25
Prerequisites   : 
- PowerShell 5.1 or higher
- Appropriate permissions in Azure AD (Files.Read.All, Sites.Read.All)
- Microsoft Graph API access

.EXAMPLE
PS> .\Get-FileswithLabels.ps1
Performs the search and exports results to a CSV file in the %TEMP% directory.

.OUTPUTS
CSV file with search results including file details and sensitivity labels.

.LINK
https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview
https://learn.microsoft.com/en-us/graph/api/driveitem-extractsensitivitylabels?view=graph-rest-1.0&tabs=http

.COMPONENT
Microsoft Graph API

.FUNCTIONALITY
- Authenticates with Microsoft Graph API using client credentials
- Performs search queries for specific files in OneDrive
- Handles pagination for large result sets
- Extracts sensitivity labels and other file metadata
- Exports results to a CSV file
#>

# ======================================
# CONFIGURATION SECTION - ADMIN SETTINGS
# ======================================
# Modify these values according to your environment

# Set your tenant name (the part before .sharepoint.com)
# Example: if your SharePoint URL is https://contoso.sharepoint.com, enter "contoso"
$tenantName = "m365cpi13246019-my"

# Set the file type to search for (without the dot)
# Common types: docx, pdf, xlsx, pptx, txt
$fileType = "docx"

# Set the tenant ID, client ID, and client secret for authentication
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3';
$clientId = 'abc64618-283f-47ba-a185-50d935d51d57';
$clientSecret = '';

# This specifies the region for the search query
$searchRegion = "NAM";

# ======================================
# CONFIGURATION SECTION - ADMIN SETTINGS
# ======================================

# Load required assemblies
Add-Type -AssemblyName System.Web

# This ensures each log file has a unique name
$date = Get-Date -Format "yyyyMMddHHmmss";

# The log file will store the search results including sensitivity labels in CSV format
$LogName = Join-Path -Path $env:TEMP -ChildPath ("OneDrive_Search_Results_" + $date + ".csv");

# Initialize global variables for the token and search results
$global:token = @();
$global:Results = @();

# This function authenticates with Microsoft Graph API and retrieves an access token
function AcquireToken() {
    # Define the URI for authentication
    $uri = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/token";

    # Define the body for the authentication request
    # Adding scope for Information Protection to access sensitivity labels
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        resource      = 'https://graph.microsoft.com'
        scope         = 'https://graph.microsoft.com/.default'
    };

    # Send the authentication request and extract the token
    $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded';
    $global:token = $loginResponse.access_token;
}


# This function sends a search request to Microsoft Graph API and handles pagination
function PerformSearch {
    # Fixing the Write-Host statement to display the search query
    Write-Host -ForegroundColor Green "Performing Search......";
    
    # Define the authorization header
    $headers = @{"Authorization" = "Bearer $global:token" };
    $string = "https://graph.microsoft.com/v1.0/search/query"; 

    # Initialize variables for pagination
    $moreresults = $true;
    $start = 0;
    $size = 200;
    $i = 0;

    # Loop to handle pagination
    while ($moreresults) {
        # The query searches for files of type 'agent' in the specified region
        $requestPayload = @"
        {
            "requests": [
                {
                    "entityTypes": [  
                    "driveItem"
                    ],
                    "query": {
                        "queryString": "(filetype:$fileType) (path:\"https://$tenantName.sharepoint.com/*\")",
                    },
                    "from": $start,
                    "size": $size,
                    "fields": [
                        "id",
                        "name",
                        "webUrl",
                        "createdDateTime",
                        "lastModifiedDateTime",
                        "createdBy",
                        "size",
                        "file",
                        "sensitivityLabel",
                        "classification",
                        "complianceLabels"
                    ],
                    "sharePointOneDriveOptions": {
                        "includeContent": "sharedContent,privateContent"
                    },
         
                     "region": "$searchRegion"
                }
            ]
        }
"@;
        # Invoke the REST method to perform the search query
        $Results = Invoke-RestMethod -Method POST -Uri $string -Headers $headers -Body $requestPayload -ContentType "application/json";

        Write-Host  $fileId 

        # Export the search results to a CSV file
        if ($null -ne $Results) {
            # Add debug output to see the structure of results
            Write-Host "Debug: Sample result structure:" -ForegroundColor Magenta
            if ($Results.value.hitsContainers.hits.Count -gt 0) {
                $sampleResult = $Results.value.hitsContainers.hits[0].resource
                Write-Host "Sample resource properties: $($sampleResult.PSObject.Properties.Name -join ', ')" -ForegroundColor Magenta
            }
            
            ExportResultSet -results $Results;
        }

        # Check if more results are available
        $moreresults = [boolean]::Parse($Results.value.hitsContainers.moreResultsAvailable);
        $start = $start + $size + 1;
        $i++;
        Write-Host -ForegroundColor Yellow "Result Batches: $i";
        Write-Host ""
    }

    Write-Host -ForegroundColor Green "Search Completed Successfully";
    Write-Host ""
    Write-Host -ForegroundColor Yellow "Results Exported to $logName";
}

# This function retrieves detailed file information including sensitivity labels
function GetFileSensitivityLabel($fileId, $driveId) {
    try {
        # Check if parameters are null
        if ($null -eq $fileId -or $null -eq $driveId) {
            Write-Warning "FileId or DriveId is null. FileId: $fileId, DriveId: $driveId"
            return "Missing IDs"
        }
        
        $headers = @{"Authorization" = "Bearer $global:token" };
        $uri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$fileId"
        
        Write-Host "Calling API: $uri" -ForegroundColor Yellow
        
        $fileDetails = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -ContentType "application/json"
        
        # Extract sensitivity label information
        $sensitivityLabel = "No Label"
        if ($fileDetails.PSObject.Properties.Name -contains 'sensitivityLabel' -and $null -ne $fileDetails.sensitivityLabel) {
            $sensitivityLabel = $fileDetails.sensitivityLabel.displayName
        }
        
        return $sensitivityLabel
    }
    catch {
        Write-Warning "Failed to retrieve sensitivity label for file ID: $fileId. Error: $($_.Exception.Message)"
        return "Error retrieving label"
    }
}

function GetSensitivityLabelViaExtractAPI($relativePath, $fileName, $resourceId = $null) {
    try {
        # With application-only tokens, we cannot use /me endpoints
        # We need to extract user information and use user-specific or site-specific endpoints
        
        # Extract user information from the relative path
        $pathParts = $relativePath.TrimStart('/').Split('/')
        if ($pathParts.Length -ge 3) {
            $siteType = $pathParts[0]  # Should be "personal" 
            $userPart = $pathParts[1]  # Should be like "willb_m365cpi13246019_onmicrosoft_com"
            $filePath = "/" + ($pathParts[2..($pathParts.Length - 1)] -join "/")
            
            Write-Host "  Parsed - Site: $siteType, User: $userPart, File: $filePath" -ForegroundColor Gray
            
            # Convert user part to proper email format for user API
            $userEmail = $userPart -replace '_', '@' -replace '@onmicrosoft@com', '.onmicrosoft.com'
            Write-Host "  Converted user email: $userEmail" -ForegroundColor Gray
            
            $headers = @{"Authorization" = "Bearer $global:token" };
            
            # If we have a resource ID, try using it with user-specific endpoints
            if ($null -ne $resourceId -and $resourceId -ne "") {
                Write-Host "  Using resource ID: $resourceId" -ForegroundColor Gray
                
                # Try Method 1: Direct user drive access with resource ID
                try {
                    $userDriveUri = "https://graph.microsoft.com/v1.0/users/$userEmail/drive"
                    Write-Host "  Getting user drive: $userDriveUri" -ForegroundColor Gray
                    
                    $userDrive = Invoke-RestMethod -Method GET -Uri $userDriveUri -Headers $headers -ContentType "application/json"
                    
                    if ($userDrive -and $userDrive.id) {
                        # Try to get file properties using the user's drive
                        $fileInfoUri = "https://graph.microsoft.com/v1.0/drives/$($userDrive.id)/items/$resourceId"
                        Write-Host "  Getting file properties: $fileInfoUri" -ForegroundColor Gray
                        
                        $fileInfo = Invoke-RestMethod -Method GET -Uri $fileInfoUri -Headers $headers -ContentType "application/json"
                        
                        if ($fileInfo) {
                            Write-Host "  File properties available: $($fileInfo.PSObject.Properties.Name -join ', ')" -ForegroundColor Gray
                            
                            # Look for sensitivity label in file properties
                            if ($fileInfo.PSObject.Properties.Name -contains 'sensitivityLabel' -and $null -ne $fileInfo.sensitivityLabel) {
                                $sensitivityLabel = $fileInfo.sensitivityLabel.displayName
                                Write-Host "  Found sensitivity label in file properties: $sensitivityLabel" -ForegroundColor Green
                                return $sensitivityLabel
                            }
                            elseif ($fileInfo.PSObject.Properties.Name -contains 'classification' -and $null -ne $fileInfo.classification) {
                                $sensitivityLabel = $fileInfo.classification
                                Write-Host "  Found classification in file properties: $sensitivityLabel" -ForegroundColor Green
                                return $sensitivityLabel
                            }
                            
                            # Try extractSensitivityLabels with user's drive
                            try {
                                $extractUri = "https://graph.microsoft.com/v1.0/drives/$($userDrive.id)/items/$resourceId/extractSensitivityLabels"
                                Write-Host "  Trying extractSensitivityLabels: $extractUri" -ForegroundColor Gray
                                
                                $extractResult = Invoke-RestMethod -Method POST -Uri $extractUri -Headers $headers -ContentType "application/json"
                                
                                # Check response structure
                                if ($extractResult -and $extractResult.labels -and $extractResult.labels.Count -gt 0) {
                                    $sensitivityLabelId = $extractResult.labels[0].sensitivityLabelId
                                    $assignmentMethod = $extractResult.labels[0].assignmentMethod
                                    $sensitivityLabel = "Label ID: $sensitivityLabelId ($assignmentMethod)"
                                    Write-Host "  Found sensitivity label via extractSensitivityLabels: $sensitivityLabel" -ForegroundColor Green
                                    return $sensitivityLabel
                                }
                                elseif ($extractResult -and $extractResult.labels -and $extractResult.labels.Count -gt 0) {
                                    $sensitivityLabelId = $extractResult.labels[0].sensitivityLabelId
                                    $assignmentMethod = $extractResult.labels[0].assignmentMethod
                                    $sensitivityLabel = "Label ID: $sensitivityLabelId ($assignmentMethod)"
                                    Write-Host "  Found sensitivity label via extractSensitivityLabels (alt): $sensitivityLabel" -ForegroundColor Green
                                    return $sensitivityLabel
                                }
                                else {
                                    Write-Host "  extractSensitivityLabels returned no labels - file has no sensitivity label applied" -ForegroundColor Yellow
                                    return "No Label"
                                }
                            }
                            catch {
                                $statusCode = $_.Exception.Response.StatusCode.value__ 
                                Write-Host "  extractSensitivityLabels failed: $statusCode - $($_.Exception.Message)" -ForegroundColor Yellow
                            }
                        }
                    }
                }
                catch {
                    Write-Host "  User drive access failed: $($_.Exception.Message)" -ForegroundColor Yellow
                }
                
                # Try Method 2: Site-based approach as fallback
                try {
                    $siteUri = "https://graph.microsoft.com/v1.0/sites/m365cpi13246019-my.sharepoint.com:/personal/$userPart"
                    Write-Host "  Trying site API: $siteUri" -ForegroundColor Gray
                
                    $siteInfo = Invoke-RestMethod -Method GET -Uri $siteUri -Headers $headers -ContentType "application/json"
                
                    if ($siteInfo -and $siteInfo.id) {
                        $driveUri = "https://graph.microsoft.com/v1.0/sites/$($siteInfo.id)/drives"
                        $drives = Invoke-RestMethod -Method GET -Uri $driveUri -Headers $headers -ContentType "application/json"
                    
                        if ($drives.value -and $drives.value.Count -gt 0) {
                            $driveId = $drives.value[0].id
                        
                            # If we have a resource ID, try using it with the site drive
                            if ($null -ne $resourceId -and $resourceId -ne "") {
                                try {
                                    $fileInfoUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$resourceId"
                                    Write-Host "  Getting file info from site drive: $fileInfoUri" -ForegroundColor Gray
                                
                                    $fileInfo = Invoke-RestMethod -Method GET -Uri $fileInfoUri -Headers $headers -ContentType "application/json"
                                
                                    if ($fileInfo) {
                                        # Look for sensitivity label in file properties
                                        if ($fileInfo.PSObject.Properties.Name -contains 'sensitivityLabel' -and $null -ne $fileInfo.sensitivityLabel) {
                                            $sensitivityLabel = $fileInfo.sensitivityLabel.displayName
                                            Write-Host "  Found sensitivity label in site file properties: $sensitivityLabel" -ForegroundColor Green
                                            return $sensitivityLabel
                                        }
                                    
                                        # Try extractSensitivityLabels
                                        try {
                                            $extractUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$resourceId/extractSensitivityLabels"
                                            Write-Host "  Trying site extractSensitivityLabels: $extractUri" -ForegroundColor Gray
                                        
                                            $extractResult = Invoke-RestMethod -Method POST -Uri $extractUri -Headers $headers -ContentType "application/json"
                                        
                                            if ($extractResult -and $extractResult.value -and $extractResult.value.labels -and $extractResult.value.labels.Count -gt 0) {
                                                $sensitivityLabelId = $extractResult.value.labels[0].sensitivityLabelId
                                                $assignmentMethod = $extractResult.value.labels[0].assignmentMethod
                                                $sensitivityLabel = "Label ID: $sensitivityLabelId ($assignmentMethod)"
                                                Write-Host "  Found sensitivity label via site extractSensitivityLabels: $sensitivityLabel" -ForegroundColor Green
                                                return $sensitivityLabel
                                            }
                                        }
                                        catch {
                                            Write-Host "  Site extractSensitivityLabels failed: $($_.Exception.Message)" -ForegroundColor Yellow
                                        }
                                    }
                                }
                                catch {
                                    Write-Host "  Site file access failed: $($_.Exception.Message)" -ForegroundColor Yellow
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-Host "  Site-based approach failed: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        
        return "No Label"
    }
    catch {
        Write-Warning "GetSensitivityLabelViaExtractAPI failed for file: $fileName. Error: $($_.Exception.Message)"
        return "API access failed"
    }
}

# This function extracts relevant fields from the search results and appends them to the CSV file
function ExportResultSet($results) {
    $Results.value.hitsContainers.hits.resource | ForEach-Object {
        Write-Host "Attempting to extract sensitivity labels using Graph API for file: $($_.name)" -ForegroundColor Cyan
        
        # Debug: Show the resource ID if available
        if ($_.PSObject.Properties.Name -contains 'id') {
            Write-Host "  Resource ID: $($_.id)" -ForegroundColor Gray
        }
        
        # Extract relative path from webUrl and try to get sensitivity label via extractSensitivityLabels API
        if ($null -ne $_.webUrl) {
            try {
                $uri = [System.Uri]$_.webUrl
                # Decode the URL
                $relativePath = [System.Web.HttpUtility]::UrlDecode($uri.AbsolutePath)
                
                Write-Host "  File path: $relativePath" -ForegroundColor Gray
                
                # Try to get the file using extractSensitivityLabels API
                # Pass the resource ID if available
                $resourceId = if ($_.PSObject.Properties.Name -contains 'id') { $_.id } else { $null }
                $sensitivityLabel = GetSensitivityLabelViaExtractAPI -relativePath $relativePath -fileName $_.name -resourceId $resourceId
            }
            catch {
                Write-Warning "Could not parse webUrl for file: $($_.name). Error: $($_.Exception.Message)"
                $sensitivityLabel = "URL parsing failed"
            }
        }
        else {
            Write-Warning "No webUrl available for file: $($_.name)"
            $sensitivityLabel = "No URL available"
        }
        
        $_ | Select-Object ID, Name, WebURL, 
        @{Name = "CreatedDate"; Expression = { $_.createdDateTime } },
        @{Name = "LastAccessedDate"; Expression = { $_.lastModifiedDateTime } },
        @{Name = "Owner"; Expression = { $_.createdBy.user.displayName } },
        @{Name = "SensitivityLabel"; Expression = { $sensitivityLabel } } | 
        Export-Csv $logName -NoTypeInformation -NoClobber -Append;
    }
}

# This is the first step before performing any search queries
AcquireToken;

# Perform search for each query
PerformSearch 
