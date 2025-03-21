<#
.SYNOPSIS
This script performs a search query on OneDrive items using Microsoft Graph API, then exports the results to a CSV file.

.DESCRIPTION
The script authenticates using OAuth with the provided tenant ID, client ID, and client secret. It reads search queries from an external file and performs a search for each query on OneDrive items. The search results are exported to a CSV file with a timestamped name.

.PARAMETER tenantId
The tenant ID for authentication.

.PARAMETER clientId
The client ID for authentication.

.PARAMETER clientSecret
The client secret for authentication.

.PARAMETER searchRegion
The region to perform the search in. Default is "NAM".

.PARAMETER searchQueryList
A list of search queries read from an external file.

.PARAMETER LogName
The path for the log file where search results will be exported.

.FUNCTION AcquireToken
Acquires an OAuth token using the provided tenant ID, client ID, and client secret.

.EXAMPLE
# Acquire OAuth token
AcquireToken;

# Read search queries from an external file
$searchQueryList = Get-Content 'C:\temp\userlist.txt'

# Perform search for each query
foreach ($searchQuery in $searchQueryList ) {
    PerformSearch -searchQuery $searchQuery;
}

.NOTES
Authors: Mike Lee, Chanchal Jain
Date: 3/21/2025

Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 

Microsoft further disclaims all implied warranties including, without limitation, 
any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, 
production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use of or inability 
to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

- Ensure that the tenant ID, client ID, and client secret are correctly set for authentication.
- Customize the ExportResultSet function as needed to select required properties and perform necessary export activities.
- The search queries should be listed in an external file (e.g., 'C:\temp\userlist.txt').
#>


# Set the tenant ID, client ID, and client secret for authentication
$tenantId = '';
$clientId = '';
$clientSecret = '';

# Generate a timestamp for the log file name
$date = Get-Date -Format "yyyyMMddHHmmss";

# Define the path for the log file
$LogName = Join-Path -Path $env:TEMP -ChildPath ("Search_Results_" + $date + ".csv");
# Set the search region
$searchRegion = "NAM";

# Initialize global variables for the token and search results
$global:token = @();
$global:Results = @();
$searchQueryList = @();

# Function to acquire an OAuth token
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
    Write-Host -ForegroundColor Green "Token acquired successfully";
}

# Function to perform a search query
function PerformSearch($searchQuery) {
    Write-Host -ForegroundColor Green "Performing Search for query: $searchQuery";
    
    # Define the authorization header
    $headers = @{"Authorization" = "Bearer $global:token" };
    $string = "https://graph.microsoft.com/v1.0/search/query"; 

    # Initialize variables for pagination
    $moreresults = $true;
    $start = 0;
    $size = 500;
    $i = 0;

    # Loop to handle pagination
    while ($moreresults) {
        # Define the request payload for the search query
        $requestPayload = @"
        {
            "requests": [
                {
                    "entityTypes": [  
                    "driveItem"
                    ],
                    "query": {
                        "queryString": "$searchQuery"
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
        Write-Host -ForegroundColor Yellow "`tLoop Counter: $i";
    }

    Write-Host -ForegroundColor Green "Search Finished Successfully";
    Write-Host ;
    Write-Host -ForegroundColor Yellow "Results Exported to $logName";
}

# Function to export the search results to a CSV file
function ExportResultSet($results) {
    Write-Host -ForegroundColor Yellow "Customize this function [ExportResultSet] as needed to select required properties and do necessary export activities";
    $Results.value.hitsContainers.hits.resource | ForEach-Object {
        $_ | Select-Object ID, Name, WebURL | Export-Csv $logName -NoTypeInformation -NoClobber -Append;
    }
}

# Acquire OAuth token
AcquireToken;

# Read search queries from an external file
$searchQueryList = Get-Content 'C:\temp\userlist.txt'

# Perform search for each query
foreach ($searchQuery in $searchQueryList ) {
    PerformSearch -searchQuery $searchQuery;
}

