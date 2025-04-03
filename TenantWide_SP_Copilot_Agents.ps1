<#
.SYNOPSIS
    Searches tenant-wide SharePoint and OneDrive locations for files of type 'agent' using Microsoft Graph API.

.DESCRIPTION
    This script authenticates to Microsoft Graph API using client credentials (tenant ID, client ID, and client secret).
    It performs a paginated search query across SharePoint and OneDrive content within the specified region (default is NAM).
    The search specifically targets files with the file type 'agent'.
    Results are exported to a CSV file stored in the user's temporary directory, with a unique timestamped filename.

.PARAMETER tenantId
    The Azure AD tenant ID used for authentication.

.PARAMETER clientId
    The client ID of the Azure AD application used for authentication.

.PARAMETER clientSecret
    The client secret associated with the Azure AD application.

.PARAMETER searchRegion
    The region to scope the search query (default is "NAM").

.OUTPUTS
    CSV file containing the following fields for each matching file:
        - ID: The unique identifier of the file.
        - Name: The name of the file.
        - WebURL: The URL to access the file.
        - CreatedDate: The date and time the file was created.
        - LastAccessedDate: The date and time the file was last modified.
        - Owner: The display name of the user who created the file.

.NOTES
    Ensure the Azure AD application has appropriate permissions to access Microsoft Graph API and search SharePoint/OneDrive content.
    The script handles pagination automatically, retrieving all available results.

    Authors: Mike Lee, Chanchal Jain
    Date: 4/3/2025

.Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 

Microsoft further disclaims all implied warranties including, without limitation, 
any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, 
production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use of or inability 
to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

    .EXAMPLE
    ./TenantWide_SP_Copilot_Agents.ps1
    Performs the search and exports results to a CSV file in the user's temporary directory.

#>

# Set the tenant ID, client ID, and client secret for authentication
$tenantId = '';
$clientId = '';
$clientSecret = '';

# This ensures each log file has a unique name
$date = Get-Date -Format "yyyyMMddHHmmss";

# The log file will store the search results in CSV format
$LogName = Join-Path -Path $env:TEMP -ChildPath ("TenantWide_SharePoint_Agents" + $date + ".csv");

# This specifies the region for the search query
$searchRegion = "NAM";

# Initialize global variables for the token and search results
$global:token = @();
$global:Results = @();

# This function authenticates with Microsoft Graph API and retrieves an access token
function AcquireToken() {
    # Define the URI for authentication
    $uri = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/token";

    # Define the body for the authentication request
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        resource      = 'https://graph.microsoft.com'
    };

    # Send the authentication request and extract the token
    $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded';
    $global:token = $loginResponse.access_token;
}

# This function sends a search request to Microsoft Graph API and handles pagination
function PerformSearch {
    # Fixing the Write-Host statement to display the search query
    Write-Host -ForegroundColor Green "Performing Search for All SharePoint Copilot Agents";
    
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
                        "queryString": "filetype:agent",
                    },
                    "from": $start,
                    "size": $size,
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


# This function extracts relevant fields from the search results and appends them to the CSV file

function ExportResultSet($results) {
    $Results.value.hitsContainers.hits.resource | ForEach-Object {
        $_ | Select-Object ID, Name, WebURL, 
        @{Name = "CreatedDate"; Expression = { $_.createdDateTime } },
        @{Name = "LastAccessedDate"; Expression = { $_.lastModifiedDateTime } },
        @{Name = "Owner"; Expression = { $_.createdBy.user.displayName } } | 
        Export-Csv $logName -NoTypeInformation -NoClobber -Append;
    }
}

# This is the first step before performing any search queries
AcquireToken;

# Perform search for each query
PerformSearch 
