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
File Name       : Get-SPOFileswithLabels.ps1
Author          : Mike Lee
Date Created    : 7/18/25
Date Updated    : 7/22/25
Prerequisites   : 
- PowerShell 5.1 or higher
- Appropriate permissions in Azure AD 

API Permissions Required:
- Sites.Read.All (for both OneDrive and SharePoint sites)
- Files.Read.All (for file access)
- InformationProtectionPolicy.Read.All (for sensitivity labels)

- Microsoft Graph API access

.EXAMPLE
PS> .\Get-SPOFileswithLabels.ps1
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
- Handles throttling using exponential backoff
- Extracts sensitivity labels and other file metadata
- Exports results to a CSV file
#>

##############################################################
#                  CONFIGURATION SECTION                    #
#############################################################
# Modify these values according to your environment

# Set your tenant name (the part before .sharepoint.com)
# Example: if your SharePoint URL is https://contoso.sharepoint.com, enter "contoso"
$tenantName = "m365cpi13246019-my"

# Set the file type to search for (without the dot)
# Common types: docx, pdf, xlsx, pptx, txt
$fileType = "docx"

# Enable or disable verbose debug output
# Set to $true for detailed logging, $false for basic info only
$debug = $false

# Set the tenant ID, client ID, and client secret for authentication
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3';
$clientId = 'abc64618-283f-47ba-a185-50d935d51d57';

# Authentication type: Choose 'ClientSecret' or 'Certificate'
$AuthType = 'Certificate'  # Valid values: 'ClientSecret' or 'Certificate'

# Batch processing configuration to avoid timeouts
$batchSize = 50  # Reduce batch size to avoid overwhelming the API
$delayBetweenBatches = 5  # Seconds to wait between batches
$maxConcurrentRequests = 5  # Limit concurrent requests

# Client Secret authentication (used when $AuthType = 'ClientSecret')
$clientSecret = '';

# Certificate authentication (used when $AuthType = 'Certificate')
$Thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"

# Certificate store location: Choose 'LocalMachine' or 'CurrentUser'
$CertStore = 'LocalMachine'  # Valid values: 'LocalMachine' or 'CurrentUser'

# This specifies the region for the search query
$searchRegion = "NAM";

#############################################################
#                  CONFIGURATION SECTION                    #
#############################################################

# Load required assemblies
Add-Type -AssemblyName System.Web

# This ensures each log file has a unique name
$date = Get-Date -Format "yyyyMMddHHmmss";

# The log file will store the search results including sensitivity labels in CSV format
$LogName = Join-Path -Path $env:TEMP -ChildPath ("SPOFileswithLabels_Search_Results_" + $date + ".csv");

# Initialize global variables for the token and search results
$global:token = @();
$global:tokenExpiry = $null;
$global:Results = @();

# Function to handle throttling for Microsoft Graph requests
# This implements best practices from https://learn.microsoft.com/en-us/graph/throttling
# It automatically handles 429 "Too Many Requests" responses with either:
# 1. The Retry-After header value if provided by the server
# 2. Exponential backoff if no Retry-After header is present
function Invoke-GraphRequestWithThrottleHandling {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [string]$Method,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers = @{},
        
        [Parameter(Mandatory = $false)]
        [string]$Body = $null,
        
        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json",
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 15,  # Increased from 10
        
        [Parameter(Mandatory = $false)]
        [int]$InitialBackoffSeconds = 3,  # Increased from 2
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 300  # 5 minute timeout
    )
    
    $retryCount = 0
    $backoffSeconds = $InitialBackoffSeconds
    $success = $false
    $result = $null
    
    if ($debug) {
        Write-Host "Making Graph request to $Uri" -ForegroundColor Gray
    }
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            # Create web request with timeout
            if ($Body) {
                $result = Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -Body $Body -ContentType $ContentType -TimeoutSec $TimeoutSeconds -ErrorAction Stop
            }
            else {
                $result = Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -ContentType $ContentType -TimeoutSec $TimeoutSeconds -ErrorAction Stop
            }
            $success = $true
        }
        catch [System.Net.WebException] {
            $webException = $_.Exception
            $statusCode = $null
            
            # Handle different types of web exceptions
            if ($webException.Response) {
                $statusCode = [int]$webException.Response.StatusCode
            }
            
            # Check for timeout or connection errors
            if ($webException.Status -eq [System.Net.WebExceptionStatus]::Timeout -or 
                $webException.Status -eq [System.Net.WebExceptionStatus]::ConnectionClosed -or
                $webException.Status -eq [System.Net.WebExceptionStatus]::ConnectFailure -or
                $statusCode -eq 502 -or $statusCode -eq 503 -or $statusCode -eq 504) {
                
                $retryCount++
                $waitTime = [Math]::Min($backoffSeconds, 300)  # Cap at 5 minutes
                
                Write-Host "Connection/timeout error detected. Status: $($webException.Status). Waiting $waitTime seconds before retry. Attempt $retryCount of $MaxRetries..." -ForegroundColor Yellow
                
                if ($retryCount -lt $MaxRetries) {
                    Start-Sleep -Seconds $waitTime
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 300)  # Exponential backoff capped at 5 minutes
                }
                else {
                    Write-Host "Maximum retry attempts reached ($MaxRetries). Giving up." -ForegroundColor Red
                    throw $_
                }
            }
            elseif ($statusCode -eq 429) {
                # Handle throttling
                $retryAfter = $null
                if ($webException.Response.Headers["Retry-After"]) {
                    $retryAfter = [int]($webException.Response.Headers["Retry-After"])
                    Write-Host "Request throttled. Retry-After header suggests waiting for $retryAfter seconds." -ForegroundColor Yellow
                }
                else {
                    $retryAfter = $backoffSeconds
                    Write-Host "Request throttled. Using exponential backoff: waiting for $retryAfter seconds." -ForegroundColor Yellow
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 300)
                }
                
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Host "Throttling detected. Waiting before retry. Attempt $retryCount of $MaxRetries..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfter
                }
                else {
                    Write-Host "Maximum retry attempts reached ($MaxRetries). Giving up." -ForegroundColor Red
                    throw $_
                }
            }
            else {
                # Not a recoverable error, rethrow
                throw $_
            }
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = $_.Exception.Response.StatusCode.value__
            }
            
            # Check if this is a throttling error (429) or server error (5xx)
            if ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -le 599)) {
                # Get the Retry-After header if it exists
                $retryAfter = $null
                if ($statusCode -eq 429 -and $_.Exception.Response.Headers.Contains("Retry-After")) {
                    $retryAfter = [int]($_.Exception.Response.Headers.GetValues("Retry-After") | Select-Object -First 1)
                    Write-Host "Request throttled. Retry-After header suggests waiting for $retryAfter seconds." -ForegroundColor Yellow
                }
                else {
                    # If no Retry-After header, use exponential backoff
                    $retryAfter = $backoffSeconds
                    Write-Host "Server error ($statusCode) or throttling detected. Using exponential backoff: waiting for $retryAfter seconds." -ForegroundColor Yellow
                    # Increase backoff for next potential retry (exponential)
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 300)  # Cap at 5 minutes
                }
                
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Host "Retryable error detected. Waiting before retry. Attempt $retryCount of $MaxRetries..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfter
                }
                else {
                    Write-Host "Maximum retry attempts reached ($MaxRetries). Giving up." -ForegroundColor Red
                    throw $_
                }
            }
            else {
                # Not a throttling error, rethrow
                throw $_
            }
        }
    }
    
    return $result
}

# This function authenticates with Microsoft Graph API and retrieves an access token
function AcquireToken() {
    Write-Host "Connecting to Microsoft Graph using $AuthType authentication..." -ForegroundColor Cyan
    
    if ($AuthType -eq 'ClientSecret') {
        # Client Secret authentication
        $uri = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/token";
        
        # Define the body for the authentication request
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $clientId
            client_secret = $clientSecret
            resource      = 'https://graph.microsoft.com'
            scope         = 'https://graph.microsoft.com/.default'
        };
        
        try {
            # Send the authentication request and extract the token
            $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop;
            $global:token = $loginResponse.access_token;
            
            # Calculate token expiry (typically 1 hour, but we'll refresh before then)
            $expiresIn = if ($loginResponse.expires_in) { $loginResponse.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)  # Refresh 5 minutes before expiry
            
            Write-Host "Successfully connected using Client Secret authentication. Token expires at: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to connect using Client Secret authentication" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }
    }
    elseif ($AuthType -eq 'Certificate') {
        # Certificate authentication
        $uri = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/v2.0/token";
        
        # Get the certificate from the local certificate store
        try {
            $cert = Get-Item -Path "Cert:\$CertStore\My\$Thumbprint" -ErrorAction Stop
        }
        catch {
            Write-Host "Certificate with thumbprint $Thumbprint not found in $CertStore\My store" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }
        
        # Create the JWT assertion for certificate authentication
        $now = [System.DateTimeOffset]::UtcNow
        $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        $nbf = $now.ToUnixTimeSeconds()
        $aud = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
        
        # Create JWT header
        $header = @{
            alg = "RS256"
            typ = "JWT"
            x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        } | ConvertTo-Json -Compress
        
        # Create JWT payload
        $payload = @{
            aud = $aud
            exp = $exp
            iss = $clientId
            jti = [System.Guid]::NewGuid().ToString()
            nbf = $nbf
            sub = $clientId
        } | ConvertTo-Json -Compress
        
        # Base64 encode header and payload
        $headerBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $payloadBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        # Create the string to sign
        $stringToSign = "$headerBase64.$payloadBase64"
        
        # Sign the string with the certificate
        $signature = $cert.PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($stringToSign), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $signatureBase64 = [Convert]::ToBase64String($signature).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        # Create the final JWT
        $jwt = "$stringToSign.$signatureBase64"
        
        # Define the body for the authentication request
        $body = @{
            client_id             = $clientId
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = $jwt
            scope                 = "https://graph.microsoft.com/.default"
            grant_type            = "client_credentials"
        }
        
        try {
            # Send the authentication request and extract the token
            $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop;
            $global:token = $loginResponse.access_token;
            
            # Calculate token expiry (typically 1 hour, but we'll refresh before then)
            $expiresIn = if ($loginResponse.expires_in) { $loginResponse.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)  # Refresh 5 minutes before expiry
            
            Write-Host "Successfully connected using Certificate authentication. Token expires at: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to connect using Certificate authentication" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }
    }
    else {
        Write-Host "Invalid authentication type: $AuthType. Valid values are 'ClientSecret' or 'Certificate'." -ForegroundColor Red
        Exit
    }
}

# Function to check if token needs refresh and refresh if necessary
function Ensure-ValidToken() {
    if ($null -eq $global:tokenExpiry -or (Get-Date) -gt $global:tokenExpiry) {
        Write-Host "Token expired or expiring soon. Refreshing..." -ForegroundColor Yellow
        AcquireToken
    }
}


# This function sends a search request to Microsoft Graph API and handles pagination
function PerformSearch {
    # Fixing the Write-Host statement to display the search query
    Write-Host -ForegroundColor Green "Performing Search......";
    
    # Define the authorization header
    $headers = @{"Authorization" = "Bearer $global:token" };
    $string = "https://graph.microsoft.com/v1.0/search/query"; 

    # Initialize variables for pagination - reduced batch size to avoid timeouts
    $moreresults = $true;
    $start = 0;
    $size = $batchSize;  # Use configurable batch size instead of hardcoded 200
    $i = 0;
    $totalProcessed = 0;

    # Loop to handle pagination
    while ($moreresults) {
        # Ensure we have a valid token before making the request
        Ensure-ValidToken
        
        Write-Host -ForegroundColor Cyan "Processing batch $($i + 1), items $($start + 1) to $($start + $size)..."
        
        # The query searches for files of type specified in the configuration
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
        # Invoke the REST method to perform the search query with enhanced error handling
        try {
            $Results = Invoke-GraphRequestWithThrottleHandling -Uri $string -Method "POST" -Headers $headers -Body $requestPayload -ContentType "application/json";
            
            # Export the search results to a CSV file
            if ($null -ne $Results) {
                # Add debug output to see the structure of results
                if ($debug) {
                    Write-Host "Debug: Sample result structure:" -ForegroundColor Magenta
                    if ($Results.value.hitsContainers.hits.Count -gt 0) {
                        $sampleResult = $Results.value.hitsContainers.hits[0].resource
                        Write-Host "Sample resource properties: $($sampleResult.PSObject.Properties.Name -join ', ')" -ForegroundColor Magenta
                    }
                }
                
                $batchItemCount = $Results.value.hitsContainers.hits.Count
                Write-Host -ForegroundColor Yellow "Processing $batchItemCount items in current batch..."
                
                # Process results in smaller chunks to avoid memory issues
                ExportResultSet -results $Results;
                
                $totalProcessed += $batchItemCount
                Write-Host -ForegroundColor Green "Total items processed so far: $totalProcessed"
            }

            # Check if more results are available
            $moreresults = [boolean]::Parse($Results.value.hitsContainers.moreResultsAvailable);
            $start = $start + $size;  # Removed the +1 which was causing gaps
            $i++;
            Write-Host -ForegroundColor Yellow "Completed batch: $i";
            
            # Add delay between batches to avoid overwhelming the API
            if ($moreresults -and $delayBetweenBatches -gt 0) {
                Write-Host -ForegroundColor Cyan "Waiting $delayBetweenBatches seconds before next batch to avoid rate limiting..."
                Start-Sleep -Seconds $delayBetweenBatches
            }
            
            Write-Host ""
        }
        catch {
            Write-Host -ForegroundColor Red "Error processing batch $($i + 1): $($_.Exception.Message)"
            
            # Determine if we should retry or stop
            if ($_.Exception.Message -like "*timeout*" -or $_.Exception.Message -like "*gateway*" -or $_.Exception.Message -like "*502*" -or $_.Exception.Message -like "*503*" -or $_.Exception.Message -like "*504*") {
                Write-Host -ForegroundColor Yellow "Network/timeout error detected. Waiting 30 seconds before continuing..."
                Start-Sleep -Seconds 30
                
                # Don't increment the batch counter, so we retry the same batch
                continue
            }
            else {
                # Non-recoverable error, rethrow
                throw $_
            }
        }
    }

    Write-Host -ForegroundColor Green "Search Completed Successfully";
    Write-Host -ForegroundColor Green "Total items processed: $totalProcessed"
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
        
        $fileDetails = Invoke-GraphRequestWithThrottleHandling -Uri $uri -Method "GET" -Headers $headers -ContentType "application/json"
        
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

function GetSensitivityLabelViaExtractAPI($relativePath, $fileName, $resourceId = $null, $webUrl = $null) {
    try {
        # Handle both OneDrive personal sites and regular SharePoint Online sites
        # OneDrive personal sites: /personal/user_domain_com/...
        # SharePoint sites: /sites/sitename/... or /teams/teamname/...
        
        # Extract site information from the relative path
        $pathParts = $relativePath.TrimStart('/').Split('/')
        if ($pathParts.Length -ge 3) {
            $siteType = $pathParts[0]  # Could be "personal", "sites", "teams", etc.
            $siteIdentifier = $pathParts[1]  # User part or site name
            $filePath = "/" + ($pathParts[2..($pathParts.Length - 1)] -join "/")
            
            if ($debug) {
                Write-Host "  Parsed - Site Type: $siteType, Identifier: $siteIdentifier, File: $filePath" -ForegroundColor Gray
            }
            
            $headers = @{"Authorization" = "Bearer $global:token" };
            
            # Handle OneDrive personal sites
            if ($siteType -eq "personal") {
                # Convert user part to proper email format for user API
                # Handle both onmicrosoft.com and vanity domains with comprehensive pattern matching
                
                if ($siteIdentifier -match '_onmicrosoft_com$') {
                    # Handle onmicrosoft.com domains
                    $userEmail = $siteIdentifier -replace '_onmicrosoft_com$', '.onmicrosoft.com'  # Fix the domain suffix first
                    $userEmail = $userEmail -replace '_([^_]+)\.onmicrosoft\.com$', '@$1.onmicrosoft.com'  # Replace the last underscore before domain with @
                    $userEmail = $userEmail -replace '_', '.'  # Convert remaining underscores to dots (for names like Will_Bob -> Will.Bob)
                }
                else {
                    # Handle all other domains with a more comprehensive approach
                    # This regex captures the domain and TLD parts more flexibly
                    # Pattern: username_parts_domain_tld where domain can contain hyphens and numbers
                    
                    # First, identify the TLD (last part after final underscore)
                    $tldMatch = $siteIdentifier -match '.*_([a-zA-Z]{2,})$'
                    if ($tldMatch) {
                        $tld = $matches[1]
                        $withoutTld = $siteIdentifier -replace "_$tld$", ""
                        
                        # Now find the domain part (everything after the last underscore in the remaining string)
                        $domainMatch = $withoutTld -match '(.*)_([a-zA-Z0-9\-]+)$'
                        if ($domainMatch) {
                            $usernamePart = $matches[1]
                            $domainPart = $matches[2]
                            
                            # Convert username underscores to dots and construct email
                            $userEmail = ($usernamePart -replace '_', '.') + "@" + $domainPart + "." + $tld
                        }
                        else {
                            # Fallback: treat the part before TLD as domain
                            $userEmail = $withoutTld -replace '_([^_]+)$', '@$1' -replace '_', '.'
                            $userEmail = $userEmail + "." + $tld
                        }
                    }
                    else {
                        # Final fallback: simple pattern replacement for cases without clear TLD
                        $userEmail = $siteIdentifier -replace '_', '@' -replace '@@', '@'
                        # Convert remaining underscores to dots after the @ symbol
                        if ($userEmail -match '@') {
                            $parts = $userEmail -split '@'
                            if ($parts.Length -eq 2) {
                                $userEmail = ($parts[0] -replace '_', '.') + '@' + ($parts[1] -replace '_', '.')
                            }
                        }
                    }
                }
                
                # Show email 
                Write-Host "   User email: $userEmail" -ForegroundColor Cyan
                
                # Try OneDrive personal site approach
                if ($null -ne $resourceId -and $resourceId -ne "") {
                    if ($debug) {
                        Write-Host "  Using OneDrive personal site approach with resource ID: $resourceId" -ForegroundColor Gray
                    }
                    
                    try {
                        $userDriveUri = "https://graph.microsoft.com/v1.0/users/$userEmail/drive"
                        if ($debug) {
                            Write-Host "  Getting user drive: $userDriveUri" -ForegroundColor Gray
                        }
                        
                        $userDrive = Invoke-GraphRequestWithThrottleHandling -Uri $userDriveUri -Method "GET" -Headers $headers -ContentType "application/json"
                        
                        if ($userDrive -and $userDrive.id) {
                            # Try to get file properties using the user's drive
                            $fileInfoUri = "https://graph.microsoft.com/v1.0/drives/$($userDrive.id)/items/$resourceId"
                            if ($debug) {
                                Write-Host "  Getting file properties: $fileInfoUri" -ForegroundColor Gray
                            }
                            
                            $fileInfo = Invoke-GraphRequestWithThrottleHandling -Uri $fileInfoUri -Method "GET" -Headers $headers -ContentType "application/json"
                            
                            if ($fileInfo) {
                                return ProcessFileForSensitivityLabel -fileInfo $fileInfo -userDrive $userDrive -resourceId $resourceId -headers $headers
                            }
                        }
                    }
                    catch {
                        if ($debug) {
                            Write-Host "  OneDrive personal site access failed: $($_.Exception.Message)" -ForegroundColor Yellow
                        }
                    }
                }
            }
            # Handle SharePoint Online sites (sites, teams, etc.)
            else {
                Write-Host "   SharePoint site: $siteIdentifier" -ForegroundColor Cyan
                
                if ($null -ne $resourceId -and $resourceId -ne "" -and $null -ne $webUrl) {
                    if ($debug) {
                        Write-Host "  Using SharePoint site approach with resource ID: $resourceId" -ForegroundColor Gray
                    }
                    
                    try {
                        # Extract the site URL from the webUrl
                        $uri = [System.Uri]$webUrl
                        $siteUrl = $uri.Scheme + "://" + $uri.Host + "/" + $pathParts[0] + "/" + $pathParts[1]
                        
                        if ($debug) {
                            Write-Host "  Site URL: $siteUrl" -ForegroundColor Gray
                        }
                        
                        # Get the site ID using the site URL 
                        $siteInfoUri = "https://graph.microsoft.com/v1.0/sites/$($uri.Host):$($uri.AbsolutePath.Split('/')[1])/$($uri.AbsolutePath.Split('/')[2])"
                        
                        # Alternative approach: use the hostname and path
                        $siteInfoUri = "https://graph.microsoft.com/v1.0/sites/$($uri.Host):/sites/$siteIdentifier"
                        
                        if ($debug) {
                            Write-Host "  Getting site info: $siteInfoUri" -ForegroundColor Gray
                        }
                        
                        $siteInfo = Invoke-GraphRequestWithThrottleHandling -Uri $siteInfoUri -Method "GET" -Headers $headers -ContentType "application/json"
                        
                        if ($siteInfo -and $siteInfo.id) {
                            # Get the default drive for the site
                            $siteDefaultDriveUri = "https://graph.microsoft.com/v1.0/sites/$($siteInfo.id)/drive"
                            if ($debug) {
                                Write-Host "  Getting site default drive: $siteDefaultDriveUri" -ForegroundColor Gray
                            }
                            
                            $siteDrive = Invoke-GraphRequestWithThrottleHandling -Uri $siteDefaultDriveUri -Method "GET" -Headers $headers -ContentType "application/json"
                            
                            if ($siteDrive -and $siteDrive.id) {
                                # Try to get file properties using the site's drive
                                $fileInfoUri = "https://graph.microsoft.com/v1.0/drives/$($siteDrive.id)/items/$resourceId"
                                if ($debug) {
                                    Write-Host "  Getting file properties: $fileInfoUri" -ForegroundColor Gray
                                }
                                
                                $fileInfo = Invoke-GraphRequestWithThrottleHandling -Uri $fileInfoUri -Method "GET" -Headers $headers -ContentType "application/json"
                                
                                if ($fileInfo) {
                                    return ProcessFileForSensitivityLabel -fileInfo $fileInfo -userDrive $siteDrive -resourceId $resourceId -headers $headers
                                }
                            }
                        }
                    }
                    catch {
                        if ($debug) {
                            Write-Host "  SharePoint site access failed: $($_.Exception.Message)" -ForegroundColor Yellow
                        }
                    }
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

# Helper function to process file for sensitivity label (common logic for both OneDrive and SharePoint)
function ProcessFileForSensitivityLabel($fileInfo, $userDrive, $resourceId, $headers) {
    if ($debug) {
        Write-Host "  File properties available: $($fileInfo.PSObject.Properties.Name -join ', ')" -ForegroundColor Gray
    }
    
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
    
    # Try extractSensitivityLabels with the drive
    try {
        $extractUri = "https://graph.microsoft.com/v1.0/drives/$($userDrive.id)/items/$resourceId/extractSensitivityLabels"
        if ($debug) {
            Write-Host "  Trying extractSensitivityLabels: $extractUri" -ForegroundColor Gray
        }
        
        $extractResult = Invoke-GraphRequestWithThrottleHandling -Uri $extractUri -Method "POST" -Headers $headers -ContentType "application/json"
        
        # Check response structure
        if ($extractResult -and $extractResult.labels -and $extractResult.labels.Count -gt 0) {
            $sensitivityLabelId = $extractResult.labels[0].sensitivityLabelId
            $assignmentMethod = $extractResult.labels[0].assignmentMethod
            $sensitivityLabel = "Label ID: $sensitivityLabelId ($assignmentMethod)"
            Write-Host "  Found sensitivity label via extractSensitivityLabels: $sensitivityLabel" -ForegroundColor Green
            return $sensitivityLabel
        }
        else {
            if ($debug) {
                Write-Host "  extractSensitivityLabels returned no labels - file has no sensitivity label applied" -ForegroundColor Yellow
            }
            return "No Label"
        }
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__ 
        if ($debug) {
            Write-Host "  extractSensitivityLabels failed: $statusCode - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        return "No Label"
    }
}

# This function extracts relevant fields from the search results and appends them to the CSV file
function ExportResultSet($results) {
    $items = $Results.value.hitsContainers.hits.resource
    $itemCount = $items.Count
    $processedCount = 0
    
    Write-Host -ForegroundColor Cyan "Processing $itemCount items for sensitivity labels..."
    
    # Process items in smaller batches to avoid overwhelming the API
    $processingBatchSize = [Math]::Min($maxConcurrentRequests, $itemCount)
    
    for ($i = 0; $i -lt $itemCount; $i += $processingBatchSize) {
        $endIndex = [Math]::Min($i + $processingBatchSize - 1, $itemCount - 1)
        $batch = $items[$i..$endIndex]
        
        Write-Host -ForegroundColor Gray "Processing items $($i + 1) to $($endIndex + 1) of $itemCount..."
        
        # Process batch items
        foreach ($item in $batch) {
            try {
                if ($debug) {
                    Write-Host "Attempting to extract sensitivity labels using Graph API for file: $($item.name)" -ForegroundColor Cyan
                }
                else {
                    Write-Host "Processing file: $($item.webUrl)" -ForegroundColor Cyan
                }
                
                # Debug: Show the resource ID if available
                if ($debug -and $item.PSObject.Properties.Name -contains 'id') {
                    Write-Host "  Resource ID: $($item.id)" -ForegroundColor Gray
                }
                
                # Extract relative path from webUrl and try to get sensitivity label via extractSensitivityLabels API
                $sensitivityLabel = "No Label"  # Default value
                
                if ($null -ne $item.webUrl) {
                    try {
                        $uri = [System.Uri]$item.webUrl
                        # Decode the URL
                        $relativePath = [System.Web.HttpUtility]::UrlDecode($uri.AbsolutePath)
                        
                        if ($debug) {
                            Write-Host "  File path: $relativePath" -ForegroundColor Gray
                        }
                        
                        # Try to get the file using extractSensitivityLabels API
                        # Pass the resource ID and webUrl if available
                        $resourceId = if ($item.PSObject.Properties.Name -contains 'id') { $item.id } else { $null }
                        $sensitivityLabel = GetSensitivityLabelViaExtractAPI -relativePath $relativePath -fileName $item.name -resourceId $resourceId -webUrl $item.webUrl
                    }
                    catch {
                        Write-Warning "Could not parse webUrl for file: $($item.name). Error: $($_.Exception.Message)"
                        $sensitivityLabel = "URL parsing failed"
                    }
                }
                else {
                    Write-Warning "No webUrl available for file: $($item.name)"
                    $sensitivityLabel = "No URL available"
                }
                
                # Export individual item to CSV
                $item | Select-Object ID, Name, WebURL, 
                @{Name = "CreatedDate"; Expression = { $_.createdDateTime } },
                @{Name = "LastAccessedDate"; Expression = { $_.lastModifiedDateTime } },
                @{Name = "Owner"; Expression = { $_.createdBy.user.displayName } },
                @{Name = "SensitivityLabel"; Expression = { $sensitivityLabel } } | 
                Export-Csv $logName -NoTypeInformation -NoClobber -Append;
                
                $processedCount++
                
                # Add small delay between individual items to avoid overwhelming the API
                if (($processedCount % 10) -eq 0) {
                    Write-Host -ForegroundColor Green "Processed $processedCount of $itemCount items..."
                    Start-Sleep -Milliseconds 500  # Small delay every 10 items
                }
            }
            catch {
                Write-Warning "Error processing file $($item.name): $($_.Exception.Message)"
                
                # Still export the item with error info
                $item | Select-Object ID, Name, WebURL, 
                @{Name = "CreatedDate"; Expression = { $_.createdDateTime } },
                @{Name = "LastAccessedDate"; Expression = { $_.lastModifiedDateTime } },
                @{Name = "Owner"; Expression = { $_.createdBy.user.displayName } },
                @{Name = "SensitivityLabel"; Expression = { "Error: $($_.Exception.Message)" } } | 
                Export-Csv $logName -NoTypeInformation -NoClobber -Append;
                
                $processedCount++
            }
        }
        
        # Add delay between processing batches
        if ($endIndex -lt ($itemCount - 1)) {
            Write-Host -ForegroundColor Yellow "Waiting 2 seconds before processing next batch..."
            Start-Sleep -Seconds 2
        }
    }
    
    Write-Host -ForegroundColor Green "Completed processing $processedCount items in this result set."
}

# This is the first step before performing any search queries
AcquireToken;

# Perform search for each query
PerformSearch 
