<#
.SYNOPSIS
    Discovers and documents all SharePoint Copilot Agents across a Microsoft 365 tenant.

.DESCRIPTION
    This script searches for SharePoint Copilot Agents across an entire Microsoft 365 tenant and 
    generates a comprehensive report with detailed information about each agent and its host site.
    The script utilizes Microsoft Graph API to search for Copilot agents and PnP PowerShell to 
    retrieve detailed site information.

    Information collected includes:
    - Copilot agent details (name, created date, last accessed date, owner)
    - Site information (template, owner, sensitivity label)
    - Security settings (information barriers, external sharing, site access restrictions)

.PARAMETER tenantname
    The name of your Microsoft 365 tenant (without .onmicrosoft.com)

.PARAMETER appID
    The Entra Application ID for authentication

.PARAMETER thumbprint
    The certificate thumbprint for authentication

.PARAMETER tenantid
    The Microsoft 365 tenant ID

.PARAMETER searchRegion
    The region for Microsoft Graph search (NAM, EMEA, APAC, etc.)

.NOTES
    File Name      : TenantWide_SP_Copilot_Agents_Insights.ps1
    Author         : Mike Lee
    Date Created   : 7/7/2025
    Prerequisite   : PnP.PowerShell module
                     Entra App with Sites.FullControl.All and Sites.Read.All permissions
                     Certificate-based authentication configured

.Disclaimer: 
The sample scripts are provided AS IS without warranty of any kind. 
Microsoft further disclaims all implied warranties including, without limitation, 
any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, 
production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use of or inability 
to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

.EXAMPLE
    .\TenantWide_SP_Copilot_Agents_Insights.ps1

.OUTPUTS
    - CSV file with all Copilot agents and their site information
    - Log file with detailed script execution information

.LINK
    https://learn.microsoft.com/microsoft-365-copilot/
#>


# ----------------------------------------------
# Set Variables
# ----------------------------------------------
$tenantname = "m365x61250205"                                   # This is your tenant name
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenantid = "85612ccb-4c28-4a34-88df-a538cc139a51"                # This is your Tenant ID
$searchRegion = "NAM"                                          # Region for Microsoft Graph search


# Script Variables
$currentDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$outputPath = "$env:TEMP\SPO_Copilot_Agents_$currentDateTime.csv"
$logPath = "$env:TEMP\SPO_Copilot_Agents_$currentDateTime.log"
$maxRetries = 5
$initialRetryDelay = 5
$adminUrl = "https://$tenantname-admin.sharepoint.com"

# Define PnP connection parameters for reuse across PnP cmdlets
$pnpConnectionParams = @{
    ClientId      = $appID
    Thumbprint    = $thumbprint
    Tenant        = $tenantid
    WarningAction = 'SilentlyContinue'
}

# Initialize global variables for the Graph token and search results
$global:graphToken = $null
$global:Results = @()

# Define the fields for the CSV file
$csvHeaders = "WebURL,Copilot Name,CreatedDate,LastAccessedDate,Owner,Site Name,Template,Site Owner,Sensitivity," + 
"Restrict Site Access Enabled,Restrict Site Discovery Enabled,External Sharing," + 
"Information Barrier Mode,Information Barrier Segments"

# Create the CSV file with headers
$csvHeaders | Out-File -FilePath $outputPath -Encoding UTF8 -Force

# Function to write log entries
Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $LogLevel = "INFO"
    )
    
    if ($LogName -ne $null) {
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : [$LogLevel] $LogEntryText" | Out-File -FilePath $LogName -Append
    }
}

# Function to authenticate with Microsoft Graph API and retrieve an access token
function Get-GraphAccessToken {
    try {
        # Get the certificate from the local certificate store using the thumbprint
        $certificate = Get-Item Cert:\LocalMachine\My\$Thumbprint -ErrorAction Stop

        # Define the URI for authentication
        $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        
        # Create the client assertion
        $jwtHeader = @{
            alg = "RS256"
            typ = "JWT"
            x5t = [System.Convert]::ToBase64String($certificate.GetCertHash())
        }
        
        $now = [DateTime]::UtcNow
        $jwtExpiry = [Math]::Floor(([DateTimeOffset]$now.AddMinutes(10)).ToUnixTimeSeconds())
        $jwtNbf = [Math]::Floor(([DateTimeOffset]$now).ToUnixTimeSeconds())
        $jwtIssueTime = [Math]::Floor(([DateTimeOffset]$now).ToUnixTimeSeconds())
        $jwtPayload = @{
            aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
            exp = $jwtExpiry
            iss = $AppId
            jti = [guid]::NewGuid().ToString()
            nbf = $jwtNbf
            sub = $AppId
            iat = $jwtIssueTime
        }
        
        # Convert to Base64
        $jwtHeaderBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($jwtHeader | ConvertTo-Json -Compress)))
        $jwtHeaderBase64 = $jwtHeaderBase64.TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        $jwtPayloadBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($jwtPayload | ConvertTo-Json -Compress)))
        $jwtPayloadBase64 = $jwtPayloadBase64.TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        # Sign the JWT
        $toSign = [System.Text.Encoding]::UTF8.GetBytes($jwtHeaderBase64 + "." + $jwtPayloadBase64)
        $rsa = $certificate.PrivateKey
        $signature = $rsa.SignData($toSign, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        
        # Convert signature to Base64
        $signatureBase64 = [System.Convert]::ToBase64String($signature)
        $signatureBase64 = $signatureBase64.TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        # Create the complete JWT token
        $jwt = $jwtHeaderBase64 + "." + $jwtPayloadBase64 + "." + $signatureBase64
        
        # Define the body for the authentication request
        $body = @{
            client_id             = $AppId
            client_assertion      = $jwt
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            scope                 = "https://graph.microsoft.com/.default"
            grant_type            = "client_credentials"
        }
        
        # Send the authentication request and extract the token
        $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded'
        $global:graphToken = $loginResponse.access_token
        Write-LogEntry -LogName $logPath -LogEntryText "Successfully authenticated with Microsoft Graph API using certificate." -LogLevel "INFO"
        Write-Host -ForegroundColor Green "Successfully authenticated with Microsoft Graph API using certificate."
        return $global:graphToken
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Authentication failed with Microsoft Graph API: $_" -LogLevel "ERROR"
        Write-Host -ForegroundColor Red "Authentication failed with Microsoft Graph API: $_"
        throw
    }
}

# Function to handle throttling with exponential backoff for PnP cmdlets
Function Invoke-PnPWithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock] $ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [string] $Operation = "PnP Operation",
        
        [Parameter(Mandatory = $false)]
        [int] $MaxRetries = 5,
        
        [Parameter(Mandatory = $false)]
        [int] $InitialRetryDelay = 5,
        
        [Parameter(Mandatory = $false)]
        [string] $LogName
    )
    
    $retryCount = 0
    $success = $false
    $result = $null
    $retryDelay = $InitialRetryDelay
    
    do {
        try {
            # Execute the provided script block
            $result = & $ScriptBlock
            $success = $true
            return $result
        }
        catch {
            $exceptionDetails = $_.Exception.ToString()
            
            # Check for common throttling-related HTTP status codes or messages
            if (($exceptionDetails -like "*429*") -or 
                ($exceptionDetails -like "*throttl*") -or 
                ($exceptionDetails -like "*too many requests*") -or
                ($exceptionDetails -like "*request limit exceeded*")) {
                
                $retryCount++
                
                # Check if maximum retries have been reached
                if ($retryCount -ge $MaxRetries) {
                    Write-LogEntry -LogName $LogName -LogEntryText "Max retries ($MaxRetries) reached for $Operation. Giving up." -LogLevel "ERROR" 
                    throw $_ # Re-throw the original exception
                }
                
                # Parse Retry-After header from the exception response if available
                $retryAfterValue = $null
                if ($_.Exception.Response -and $_.Exception.Response.Headers -and $_.Exception.Response.Headers["Retry-After"]) {
                    $retryAfterValue = [int]$_.Exception.Response.Headers["Retry-After"]
                    $retryDelay = $retryAfterValue # Use server-suggested delay
                    Write-LogEntry -LogName $LogName -LogEntryText "Throttling detected for $Operation. Server requested retry after $retryAfterValue seconds." -LogLevel "WARNING"
                }
                else {
                    # Use exponential backoff if no Retry-After header is present
                    $retryDelay = [Math]::Min(60, $retryDelay * 2) # Double the delay, max 60 seconds
                    Write-LogEntry -LogName $LogName -LogEntryText "Throttling detected for $Operation. Using exponential backoff: waiting $retryDelay seconds before retry $retryCount of $MaxRetries." -LogLevel "WARNING"
                }
                
                Write-Host "Throttling detected for $Operation. Waiting $retryDelay seconds before retry $retryCount of $MaxRetries." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryDelay # Wait before retrying
            }
            else {
                # If not a throttling error, re-throw the original exception
                throw $_
            }
        }
    } while (-not $success -and $retryCount -lt $MaxRetries)
}

# Function to get SharePoint site information using PnP PowerShell
Function Get-PnPSiteInformation {
    param (
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string] $CopilotName = "",
        
        [Parameter(Mandatory = $false)]
        [DateTime] $CreatedDate = $null,
        
        [Parameter(Mandatory = $false)]
        [DateTime] $LastAccessedDate = $null,
        
        [Parameter(Mandatory = $false)]
        [string] $Owner = ""
    )
    
    try {
        # Create custom object to store site information
        $siteInfo = [PSCustomObject]@{
            "WebURL"                          = $SiteUrl
            "Copilot Name"                    = $CopilotName
            "CreatedDate"                     = $CreatedDate
            "LastAccessedDate"                = $LastAccessedDate
            "Owner"                           = $Owner
            "Site Name"                       = ""
            "Template"                        = ""
            "Site Owner"                      = ""
            "Sensitivity"                     = ""
            "Restrict Site Access Enabled"    = ""
            "Restrict Site Discovery Enabled" = ""
            "External Sharing"                = ""
            "Information Barrier Mode"        = ""
            "Information Barrier Segments"    = ""
        }
        
        # Validate site URL format
        if ($SiteUrl.EndsWith("/")) {
            $SiteUrl = $SiteUrl.TrimEnd('/')
            Write-LogEntry -LogName $logPath -LogEntryText "Removed trailing slash from site URL: $SiteUrl" -LogLevel "INFO"
        }
        
        # Connect to the site using PnP
        Write-LogEntry -LogName $logPath -LogEntryText "Connecting to site: $SiteUrl" -LogLevel "INFO"
        
        # First connect to the admin center to get tenant-level site properties
        Invoke-PnPWithRetry -ScriptBlock { 
            Connect-PnPOnline -Url $adminUrl @pnpConnectionParams -ErrorAction Stop 
        } -Operation "Connect to SharePoint Admin Center" -LogName $logPath
        
        # Get tenant site properties consistently for all sites (SharePoint and OneDrive)
        $siteProp = $null
        
        try {
            # Try to get the site properties from the tenant admin
            $siteProp = Invoke-PnPWithRetry -ScriptBlock { 
                Get-PnPTenantSite -Identity $SiteUrl -ErrorAction Stop | 
                Select-Object Url, Title, Owner, Template, SensitivityLabel, 
                RestrictedAccessControl, RestrictContentOrgWideSearch, SharingCapability, 
                InformationBarrierMode, InformationBarrierSegments, CreatedDate, LastContentModifiedDate
            } -Operation "Get-PnPTenantSite for $SiteUrl" -LogName $logPath
        }
        catch {
            Write-LogEntry -LogName $logPath -LogEntryText "Error getting tenant site information for $SiteUrl : $_" -LogLevel "WARNING"
            # Try connecting directly to the site to get basic information
            try {
                Invoke-PnPWithRetry -ScriptBlock { 
                    Connect-PnPOnline -Url $SiteUrl @pnpConnectionParams -ErrorAction Stop 
                } -Operation "Connect directly to site $SiteUrl" -LogName $logPath
                
                $web = Invoke-PnPWithRetry -ScriptBlock { 
                    Get-PnPWeb -Includes Title, Created, LastItemModifiedDate, WebTemplate
                } -Operation "Get-PnPWeb for $SiteUrl" -LogName $logPath
                
                # Extract username from OneDrive URL for better title if applicable
                $userAlias = ""
                $isOneDrive = $false
                if ($SiteUrl -like "*-my.sharepoint.com*" -and $SiteUrl -match "/personal/([^/]+)") {
                    $userAlias = $Matches[1]
                    $isOneDrive = $true
                }
                
                # Create a site properties object with null values for restriction properties
                # This ensures we don't default to any specific value until we know what the actual value is
                $siteProp = [PSCustomObject]@{
                    Url                          = $SiteUrl
                    Title                        = if ($web.Title) { $web.Title } elseif ($isOneDrive) { "OneDrive for $userAlias" } else { $SiteUrl }
                    Owner                        = $Owner
                    Template                     = if ($web.WebTemplate) { $web.WebTemplate } elseif ($isOneDrive) { "SPSPERS" } else { "" }
                    SensitivityLabel             = ""
                    RestrictedAccessControl      = $null  
                    RestrictContentOrgWideSearch = $null
                    SharingCapability            = ""
                    InformationBarrierMode       = ""
                    InformationBarrierSegments   = ""
                    CreatedDate                  = $web.Created
                    LastContentModifiedDate      = $web.LastItemModifiedDate
                }
            }
            catch {
                Write-LogEntry -LogName $logPath -LogEntryText "Error connecting directly to site $SiteUrl : $_" -LogLevel "WARNING"
            }
        }
        
        if ($siteProp) {
            # Update siteInfo object with tenant site properties
            $siteInfo."WebURL" = $siteProp.Url
            $siteInfo."Site Name" = $siteProp.Title
            $siteInfo."Site Owner" = $siteProp.Owner
            
            # Only update owner if it wasn't provided
            if ([string]::IsNullOrEmpty($Owner)) {
                $siteInfo."Owner" = $siteProp.Owner
            }
            
            $siteInfo."Template" = $siteProp.Template
            
            # Get the sensitivity label display name and GUID
            if (-not [string]::IsNullOrEmpty($siteProp.SensitivityLabel)) {
                $siteInfo."Sensitivity" = Get-SensitivityLabelDisplayName -LabelId $siteProp.SensitivityLabel
            }
            else {
                $siteInfo."Sensitivity" = "Not Set"
            }
            
            # Handle RestrictedAccessControl correctly
            if ($null -ne $siteProp.RestrictedAccessControl) {
                $siteInfo."Restrict Site Access Enabled" = $siteProp.RestrictedAccessControl
            }
            else {
                # Only default to False if we couldn't determine the actual value
                $siteInfo."Restrict Site Access Enabled" = $false
                Write-LogEntry -LogName $logPath -LogEntryText "Could not determine RestrictedAccessControl for $SiteUrl, defaulting to False" -LogLevel "WARNING"
            }
            
            # Handle RestrictContentOrgWideSearch correctly
            if ($null -ne $siteProp.RestrictContentOrgWideSearch) {
                $siteInfo."Restrict Site Discovery Enabled" = $siteProp.RestrictContentOrgWideSearch
            }
            else {
                # Only default to False if we couldn't determine the actual value
                $siteInfo."Restrict Site Discovery Enabled" = $false
                Write-LogEntry -LogName $logPath -LogEntryText "Could not determine RestrictContentOrgWideSearch for $SiteUrl, defaulting to False" -LogLevel "WARNING"
            }
            
            $siteInfo."External Sharing" = $siteProp.SharingCapability
            $siteInfo."Information Barrier Mode" = $siteProp.InformationBarrierMode
            $siteInfo."Information Barrier Segments" = $siteProp.InformationBarrierSegments
            
            # Only update dates if they weren't provided
            if ($null -eq $CreatedDate) {
                $siteInfo."CreatedDate" = $siteProp.CreatedDate
            }
            
            if ($null -eq $LastAccessedDate) {
                $siteInfo."LastAccessedDate" = $siteProp.LastContentModifiedDate
            }
        }
        
        # Now connect to the specific site to get site-level information
        Invoke-PnPWithRetry -ScriptBlock { 
            Connect-PnPOnline -Url $SiteUrl @pnpConnectionParams -ErrorAction Stop 
        } -Operation "Connect to site $SiteUrl" -LogName $logPath
        
        # Get site collection administrators if not already set
        if ([string]::IsNullOrEmpty($siteInfo."Site Owner")) {
            $siteAdmins = Invoke-PnPWithRetry -ScriptBlock { 
                Get-PnPSiteCollectionAdmin 
            } -Operation "Get-PnPSiteCollectionAdmin for $SiteUrl" -LogName $logPath
            
            if ($siteAdmins -and $siteAdmins.Count -gt 0) {
                $primaryAdmin = $siteAdmins | Select-Object -First 1
                $adminName = $primaryAdmin.Title
                $adminEmail = $primaryAdmin.Email
                
                if ($adminName -and $adminEmail) {
                    $siteInfo."Site Owner" = "$adminName <$adminEmail>"
                }
                elseif ($adminName) {
                    $siteInfo."Site Owner" = $adminName
                }
            }
        }
        
        return $siteInfo
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error getting site information for $SiteUrl : $_" -LogLevel "ERROR"
        throw $_
    }
    finally {
        # Disconnect from the site
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore any disconnection errors
        }
    }
}

# Function to search for Copilot agents using Microsoft Graph API
function Search-CopilotAgents {
    # Display search query information
    Write-Host -ForegroundColor Green "Performing Search for All SharePoint Copilot Agents";
    Write-LogEntry -LogName $logPath -LogEntryText "Performing Search for All SharePoint Copilot Agents" -LogLevel "INFO"
    
    # Ensure we have a valid token
    if (-not $global:graphToken) {
        Get-GraphAccessToken
    }
    
    # Define the authorization header
    $headers = @{"Authorization" = "Bearer $global:graphToken" };
    $graphUrl = "https://graph.microsoft.com/v1.0/search/query"; 

    # Initialize variables for pagination
    $moreresults = $true;
    $start = 0;
    $size = 200;
    $i = 0;
    $agentCount = 0;
    $sites = @{}  # Hashtable to store unique sites

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
        try {
            # Invoke the REST method to perform the search query
            $searchResults = Invoke-RestMethod -Method POST -Uri $graphUrl -Headers $headers -Body $requestPayload -ContentType "application/json";

            # Process search results
            if ($null -ne $searchResults) {
                Write-Host -ForegroundColor Cyan "Processing batch of search results..."
                
                # Process each hit
                $hits = $searchResults.value.hitsContainers.hits
                if ($hits) {
                    $agentCount += $hits.Count
                    
                    foreach ($hit in $hits) {
                        $resource = $hit.resource
                        
                        if ($resource) {
                            # Extract site URL from webUrl
                            $webUrl = $resource.webUrl
                            
                            # Get the site URL using the helper function
                            $siteUrl = Get-SiteUrlFromWebUrl -WebUrl $webUrl
                            
                            # Only process if we have a valid site URL
                            if (-not [string]::IsNullOrEmpty($siteUrl)) {
                                # Store unique sites with their agent information
                                if (-not $sites.ContainsKey($siteUrl)) {
                                    $sites[$siteUrl] = @()
                                }
                            
                                # Add agent information to the site
                                $agentInfo = @{
                                    Name                 = $resource.name
                                    WebUrl               = $webUrl
                                    CreatedDateTime      = if ($resource.createdDateTime) { [DateTime]$resource.createdDateTime } else { $null }
                                    LastModifiedDateTime = if ($resource.lastModifiedDateTime) { [DateTime]$resource.lastModifiedDateTime } else { $null }
                                    CreatedBy            = if ($resource.createdBy.user.displayName) { $resource.createdBy.user.displayName } else { "" }
                                }
                                
                                $sites[$siteUrl] += $agentInfo
                            }
                            else {
                                Write-LogEntry -LogName $logPath -LogEntryText "Could not determine valid site URL from: $webUrl" -LogLevel "WARNING"
                            }
                        }
                    }
                }
            }

            # Check if more results are available
            $moreresults = [boolean]::Parse($searchResults.value.hitsContainers.moreResultsAvailable);
            $start = $start + $size;
            $i++;
            Write-Host -ForegroundColor Yellow "Result Batches: $i | Agents found: $agentCount | Unique sites: $($sites.Count)";
            Write-LogEntry -LogName $logPath -LogEntryText "Processed batch $i : Found $agentCount agents across $($sites.Count) unique sites so far." -LogLevel "INFO"
        }
        catch {
            Write-LogEntry -LogName $logPath -LogEntryText "Error searching for Copilot agents: $_" -LogLevel "ERROR"
            Write-Host -ForegroundColor Red "Error searching for Copilot agents: $_"
            
            # Check if token expired and refresh if needed
            if ($_.Exception.Message -like "*token has expired*" -or $_.Exception.Message -like "*Authentication failed*") {
                Write-Host -ForegroundColor Yellow "Token expired. Refreshing token and retrying..."
                Write-LogEntry -LogName $logPath -LogEntryText "Token expired. Refreshing token and retrying..." -LogLevel "WARNING"
                Get-GraphAccessToken
                $headers = @{"Authorization" = "Bearer $global:graphToken" };
            }
            else {
                throw $_
            }
        }
    }

    Write-Host -ForegroundColor Green "Search Completed Successfully. Found $agentCount Copilot agents across $($sites.Count) sites.";
    Write-LogEntry -LogName $logPath -LogEntryText "Search completed successfully. Found $agentCount Copilot agents across $($sites.Count) sites." -LogLevel "INFO"
    
    return $sites
}

# Function to handle site URL parsing from web URL
function Get-SiteUrlFromWebUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string] $WebUrl
    )
    
    try {
        $uri = [System.Uri]$WebUrl
        $siteUrl = ""
        
        # Check if it's a OneDrive site (contains "-my.sharepoint.com")
        if ($uri.Host -like "*-my.sharepoint.com") {
            # For OneDrive personal sites, we need to include the personal/{username} part
            if ($uri.AbsolutePath -match "/personal/([^/]+)") {
                $personalSite = $Matches[1]
                $siteUrl = "$($uri.Scheme)://$($uri.Host)/personal/$personalSite"
            }
            else {
                # If we can't extract the personal site, use the parent URL
                $siteUrl = "$($uri.Scheme)://$($uri.Host)"
            }
        }
        else {
            # For SharePoint team sites, we need to include the full site path
            # Look for "sites/" or "teams/" in the path to identify site collection
            if ($uri.AbsolutePath -match "/(sites|teams)/([^/]+)") {
                $siteType = $Matches[1]  # "sites" or "teams"
                $siteName = $Matches[2]  # The site name
                $siteUrl = "$($uri.Scheme)://$($uri.Host)/$siteType/$siteName"
            }
            else {
                # If we can't identify a subsite, use the root site
                $siteUrl = "$($uri.Scheme)://$($uri.Host)"
            }
        }
        
        return $siteUrl
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error parsing site URL from $WebUrl : $_" -LogLevel "ERROR"
        return ""
    }
}

# Function to check if a module is installed and install if not
Function Test-AndInstallModule {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ModuleName
    )
    
    try {
        if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
            Write-LogEntry -LogName $logPath -LogEntryText "Module $ModuleName is not installed. Attempting to install..." -LogLevel "INFO"
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
            Write-LogEntry -LogName $logPath -LogEntryText "Module $ModuleName installed successfully." -LogLevel "INFO"
        }
        else {
            Write-LogEntry -LogName $logPath -LogEntryText "Module $ModuleName is already installed." -LogLevel "INFO"
        }
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error installing module $ModuleName : $_" -LogLevel "ERROR"
        throw $_
    }
}

# Function to get sensitivity label display name from GUID
Function Get-SensitivityLabelDisplayName {
    param (
        [Parameter(Mandatory = $true)]
        [string] $LabelId
    )
    
    try {
        # Define a mapping of known sensitivity label GUIDs to display names
        # Note: In a real environment, populate this with actual sensitivity labels from your tenant
        $labelMapping = @{
            # Example mappings - replace with actual values from your tenant
            "e42fd39b-8d6f-43fb-b8f1-b5580df51bc0" = "Confidential"
            "2fc7f16d-67f5-46a7-b6f2-6f9c3b8269c0" = "Highly Confidential"
            "27451a5b-5823-4853-bcd4-2204d03ab477" = "Public"
            "da5e35a2-4496-4663-9c5f-1b7467b5c835" = "Internal"
            # Add more mappings as needed
        }
        
        # Check if the label ID exists in our mapping
        if ($labelMapping.ContainsKey($LabelId)) {
            $displayName = $labelMapping[$LabelId]
            return "$displayName ($LabelId)"
        }
        else {
            # If not found in our mapping, just return the GUID
            return $LabelId
        }
    }
    catch {
        Write-LogEntry -LogName $logPath -LogEntryText "Error getting sensitivity label display name for $LabelId : $_" -LogLevel "WARNING"
        return $LabelId
    }
}

# Main script execution
try {
    # Create log file
    "" | Out-File -FilePath $logPath -Force
    Write-LogEntry -LogName $logPath -LogEntryText "Script started. Collecting SharePoint Copilot Agents information." -LogLevel "INFO"
    
    # Display banner with tenant information
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  SharePoint Online Copilot Agents Information Collector" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "Tenant Name: $tenantName" -ForegroundColor Yellow
    Write-Host "Admin URL: $adminUrl" -ForegroundColor Yellow
    Write-Host "Tenant ID: $TenantId" -ForegroundColor Yellow
    Write-Host "Application ID: $AppId" -ForegroundColor Yellow
    Write-Host "Search Region: $searchRegion" -ForegroundColor Yellow
    Write-Host "Output File: $outputPath" -ForegroundColor Yellow
    Write-Host "Log File: $logPath" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Cyan
    
    # Check for PnP.PowerShell module
    Test-AndInstallModule -ModuleName "PnP.PowerShell"
    
    # Get Microsoft Graph token
    Get-GraphAccessToken
    
    # Search for Copilot agents
    $sites = Search-CopilotAgents
    
    if ($sites.Count -eq 0) {
        Write-Host "No Copilot agents found in the tenant." -ForegroundColor Yellow
        Write-LogEntry -LogName $logPath -LogEntryText "No Copilot agents found in the tenant." -LogLevel "WARNING"
    }
    else {
        # Process each site to get detailed information
        $processedCount = 0
        $totalSites = $sites.Count
        
        foreach ($siteUrl in $sites.Keys) {
            $processedCount++
            
            Write-Host "Processing site $processedCount/$totalSites : $siteUrl" -ForegroundColor Cyan
            Write-LogEntry -LogName $logPath -LogEntryText "Processing site $processedCount/$totalSites : $siteUrl" -LogLevel "INFO"
            
            try {
                # Get the agent information for this site
                $agents = $sites[$siteUrl]
                
                foreach ($agent in $agents) {
                    # Get site information using PnP
                    $siteInfo = Get-PnPSiteInformation -SiteUrl $siteUrl -CopilotName $agent.Name -CreatedDate $agent.CreatedDateTime -LastAccessedDate $agent.LastModifiedDateTime -Owner $agent.CreatedBy
                    
                    # Create a custom object for proper CSV export
                    $ibSegments = if ($siteInfo."Information Barrier Segments" -is [array]) {
                        $siteInfo."Information Barrier Segments" -join ";"
                    }
                    else {
                        $siteInfo."Information Barrier Segments"
                    }
                    
                    $exportObject = [PSCustomObject]@{
                        "WebURL"                          = $siteInfo."WebURL"
                        "Copilot Name"                    = $siteInfo."Copilot Name"
                        "CreatedDate"                     = $siteInfo."CreatedDate"
                        "LastAccessedDate"                = $siteInfo."LastAccessedDate"
                        "Owner"                           = $siteInfo."Owner"
                        "Site Name"                       = $siteInfo."Site Name"
                        "Template"                        = $siteInfo."Template"
                        "Site Owner"                      = $siteInfo."Site Owner"
                        "Sensitivity"                     = $siteInfo."Sensitivity"
                        "Restrict Site Access Enabled"    = $siteInfo."Restrict Site Access Enabled"
                        "Restrict Site Discovery Enabled" = $siteInfo."Restrict Site Discovery Enabled"
                        "External Sharing"                = $siteInfo."External Sharing"
                        "Information Barrier Mode"        = $siteInfo."Information Barrier Mode"
                        "Information Barrier Segments"    = $ibSegments
                    }
                    
                    # Export to CSV using Export-Csv which handles CSV formatting correctly
                    $exportObject | Export-Csv -Path $outputPath -NoTypeInformation -Append -Encoding UTF8
                    Write-LogEntry -LogName $logPath -LogEntryText "Successfully exported data for agent '$($agent.Name)' at site: $siteUrl" -LogLevel "INFO"
                }
            }
            catch {
                Write-LogEntry -LogName $logPath -LogEntryText "Error processing site $siteUrl : $_" -LogLevel "ERROR"
                Write-Host "Error processing site $siteUrl : $_" -ForegroundColor Red
            }
        }
        
        Write-Host "Script completed. Processed $processedCount out of $totalSites sites." -ForegroundColor Green
        Write-LogEntry -LogName $logPath -LogEntryText "Script completed. Processed $processedCount out of $totalSites sites." -LogLevel "INFO"
    }
    
    Write-Host "Output CSV file: $outputPath" -ForegroundColor Green
    Write-Host "Log file: $logPath" -ForegroundColor Green
}
catch {
    Write-Host "Error executing script: $_" -ForegroundColor Red
    Write-LogEntry -LogName $logPath -LogEntryText "Error executing script: $_" -LogLevel "ERROR"
}
finally {
    # Disconnect from SharePoint Online
    try {
        if (Get-PnPConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-LogEntry -LogName $logPath -LogEntryText "Disconnected from PnP Online." -LogLevel "INFO"
        }
    }
    catch {
        # Ignore any disconnection errors
    }
}
