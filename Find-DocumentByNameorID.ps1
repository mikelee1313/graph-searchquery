<#
.SYNOPSIS
    Finds a SharePoint Online document by GUID or name across the tenant.

.DESCRIPTION
    Presents a menu to search by:
      1. Document Name  - tenant-wide Graph API keyword search
      2. Document GUID  - exact match via Graph API

    Uses app-only certificate authentication directly against Microsoft Graph.
    No PnP PowerShell module required.

    Results are printed to the console.

.NOTES
    Authors: Mike Lee
    Updated: 5/6/2026

    Requires an Entra app registration with:
      - Graph: Sites.Read.All, Files.Read.All
    Requires a certificate (with private key) installed in the Windows cert store.

.Disclaimer: The sample scripts are provided AS IS without warranty of any kind.
    Microsoft further disclaims all implied warranties including, without limitation,
    any implied warranties of merchantability or of fitness for a particular purpose.
    The entire risk arising out of the use or performance of the sample scripts and
    documentation remains with you.
#>

#region Configuration
# ----------------------------------------------
# Set Variables
# ----------------------------------------------
$appID = "abc64618-283f-47ba-a185-50d935d51d57"     # Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9" # Certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"     # Tenant ID (GUID)
$CertStore = "LocalMachine"   # Certificate store: 'CurrentUser' or 'LocalMachine'
$searchRegion = ""             # Graph Search region (leave empty to auto-detect, or set: NAM, EUR, APC, GBR, CAN, ...)
$debugLogging = $false         # Set to $true for verbose DEBUG output

# ----------------------------------------------
# Initialize
# ----------------------------------------------
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$log  = if ($debugLogging) { "$env:TEMP\Find-DocByID_$date.log" } else { $null }

$global:token = $null
$global:tokenExpiry = $null
#endregion Configuration

#region Logging
# ----------------------------------------------
# Logging
# ----------------------------------------------
Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $Level = "INFO"
    )
    if ($Level -eq "ERROR" -or $Level -eq "INFO" -or ($Level -eq "DEBUG" -and $debugLogging)) {
        if ($null -ne $LogName) {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path $LogName -Value "[$timestamp] [$Level] $LogEntryText"
        }
    }
}

Function Write-InfoLog { param([string]$LogName, [string]$LogEntryText) Write-LogEntry -LogName $LogName -LogEntryText $LogEntryText -Level "INFO" }
Function Write-DebugLog { param([string]$LogName, [string]$LogEntryText) Write-LogEntry -LogName $LogName -LogEntryText $LogEntryText -Level "DEBUG" }
Function Write-ErrorLog { param([string]$LogName, [string]$LogEntryText) Write-LogEntry -LogName $LogName -LogEntryText $LogEntryText -Level "ERROR" }
#endregion Logging

#region Throttle Handling
# ----------------------------------------------
# Throttle Handling
# ----------------------------------------------
Function Invoke-WithThrottleHandling {
    param (
        [Parameter(Mandatory)] [scriptblock] $ScriptBlock,
        [int]    $MaxRetries = 5,
        [string] $Operation = "Operation"
    )
    $retryCount = 0
    while ($retryCount -le $MaxRetries) {
        try {
            return (& $ScriptBlock)
        }
        catch {
            $msg = $_.Exception.Message
            $isThrottling = $false
            $waitTime = 10

            if ($null -ne $_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                if ($statusCode -in @(429, 503)) {
                    $isThrottling = $true
                    $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                    if ($retryAfter) { $waitTime = [int]$retryAfter }
                }
            }
            elseif ($msg -match "throttl|Too many requests|429|503|Request limit exceeded") {
                $isThrottling = $true
                if ($msg -match "retry after (\d+)") { $waitTime = [int]$Matches[1] }
            }

            if ($isThrottling -and $retryCount -lt $MaxRetries) {
                Write-Host "  Throttled during '$Operation'. Waiting $waitTime seconds (retry $($retryCount+1)/$MaxRetries)..." -ForegroundColor Yellow
                Write-InfoLog -LogName $log -LogEntryText "Throttled during '$Operation'. Waiting ${waitTime}s (retry $($retryCount+1)/$MaxRetries)"
                Start-Sleep -Seconds $waitTime
                $retryCount++
                continue
            }

            throw
        }
    }
}
#endregion Throttle Handling

#region Authentication
# ----------------------------------------------
# Certificate-Based Graph Authentication
# ----------------------------------------------
Function AcquireToken {
    Write-Host "  Authenticating to Microsoft Graph (Certificate)..." -ForegroundColor Cyan
    $tokenUri = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token"

    try {
        $cert = Get-Item -Path "Cert:\$CertStore\My\$thumbprint" -ErrorAction Stop
    }
    catch {
        Write-Host "  Certificate $thumbprint not found in Cert:\$CertStore\My" -ForegroundColor Red
        throw
    }

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    if (-not $rsa) { throw "Unable to access RSA private key for certificate $thumbprint." }

    $now = [System.DateTimeOffset]::UtcNow
    $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
    $nbf = $now.ToUnixTimeSeconds()
    $x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $header = @{ alg = 'RS256'; typ = 'JWT'; x5t = $x5t } | ConvertTo-Json -Compress
    $payload = @{
        aud = $tokenUri
        exp = $exp
        iss = $appID
        jti = [System.Guid]::NewGuid().ToString()
        nbf = $nbf
        sub = $appID
    } | ConvertTo-Json -Compress

    $hB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $pB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $toSign = "$hB64.$pB64"
    $sig = $rsa.SignData(
        [System.Text.Encoding]::UTF8.GetBytes($toSign),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $jwt = "$toSign.$([Convert]::ToBase64String($sig).TrimEnd('=').Replace('+','-').Replace('/','_'))"

    $body = @{
        client_id             = $appID
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = $jwt
        scope                 = 'https://graph.microsoft.com/.default'
        grant_type            = 'client_credentials'
    }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
        $global:token = $resp.access_token
        $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
        $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
        Write-Host "  Authenticated via Certificate ($thumbprint). Token valid until: $($global:tokenExpiry)" -ForegroundColor Green
        Write-InfoLog -LogName $log -LogEntryText "Authenticated via certificate. Token valid until $($global:tokenExpiry)"
    }
    catch {
        Write-Host "  Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

Function Test-ValidToken {
    if ($null -eq $global:tokenExpiry -or (Get-Date) -gt $global:tokenExpiry) {
        Write-Host '  Token expired or missing - refreshing...' -ForegroundColor Yellow
        AcquireToken
    }
}

Function Get-GraphAuthHeaders {
    Test-ValidToken
    return @{
        'Authorization' = "Bearer $global:token"
        'Content-Type'  = 'application/json'
    }
}
#endregion Authentication

#region Region Detection
# ----------------------------------------------
# Graph Search Region Detection
# ----------------------------------------------
Function Test-GraphSearchRegion {
    param([string] $Region, [hashtable] $Headers)

    $testQuery = @{
        requests = @(@{
                entityTypes = @("driveItem")
                query       = @{ queryString = "test" }
                from = 0; size = 1
                region      = $Region
            })
    }
    try {
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/search/query" `
            -Headers $Headers -Method Post `
            -Body ($testQuery | ConvertTo-Json -Depth 5) `
            -ErrorAction Stop | Out-Null
        return $true
    }
    catch { return $false }
}

Function Get-GraphSearchRegion {
    Write-Host "  Auto-detecting Microsoft Graph Search region..." -ForegroundColor Cyan

    $headers = Get-GraphAuthHeaders

    $regionsToProbe = @(
        "NAM", "US", "EUR", "APC", "GBR", "CAN", "IND", "AUS", "JPN", "DEU",
        "ZAF", "ARE", "CHE", "NOR", "KOR", "SWE", "TWN", "FRA", "ITA", "MEX", "LAM"
    )

    foreach ($region in $regionsToProbe) {
        if (Test-GraphSearchRegion -Region $region -Headers $headers) {
            Write-Host "  Detected Graph Search region: $region" -ForegroundColor Green
            Write-InfoLog -LogName $log -LogEntryText "Auto-detected Graph Search region: $region"
            return $region
        }
    }

    Write-Host "  Could not auto-detect region. Defaulting to NAM." -ForegroundColor Yellow
    return "NAM"
}
#endregion Region Detection

#region Search Functions
# ----------------------------------------------
# Search for a document via Microsoft Graph API (tenant-wide)
# Returns all matching hits, not just the first one.
# ----------------------------------------------
Function Search-DocumentViaGraphAPI {
    param(
        [Parameter(Mandatory)] [string] $DocumentId,
        [Parameter(Mandatory)] [string] $SearchRegion
    )

    $results = [System.Collections.Generic.List[PSObject]]::new()

    try {
        $headers = Get-GraphAuthHeaders
        $searchUrl = "https://graph.microsoft.com/v1.0/search/query"

        foreach ($entityType in @("driveItem", "listItem")) {
            $body = @{
                requests = @(@{
                        entityTypes               = @($entityType)
                        query                     = @{ queryString = "`"$DocumentId`"" }
                        from                      = 0
                        size                      = 25
                        sharePointOneDriveOptions = @{ includeContent = "sharedContent,privateContent" }
                        region                    = $SearchRegion
                    })
            } | ConvertTo-Json -Depth 5

            $response = Invoke-WithThrottleHandling -ScriptBlock {
                Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body $body
            } -Operation "Graph search ($entityType) for $DocumentId"

            $hits = $response.value[0].hitsContainers[0].hits
            if ($hits -and $hits.Count -gt 0) {
                foreach ($hit in $hits) {
                    $r = $hit.resource
                    if (-not $r) { continue }

                    $owner = if ($r.createdBy.user.displayName) {
                        if ($r.createdBy.user.email) {
                            "$($r.createdBy.user.displayName) <$($r.createdBy.user.email)>"
                        }
                        else { $r.createdBy.user.displayName }
                    }
                    else { "Unknown" }

                    # driveItem has .name; listItem does not — fall back to the URL leaf
                    $itemName = if ($r.PSObject.Properties['name'] -and $r.name) {
                        $r.name
                    }
                    elseif ($r.webUrl) {
                        try { Split-Path ([System.Uri]::new($r.webUrl).LocalPath) -Leaf } catch { $r.webUrl }
                    }
                    else { "Unknown" }

                    # SharePoint UniqueId is in sharepointIds.listItemUniqueId;
                    # for listItem results, $r.id is also the item GUID
                    $uniqueId = if ($r.PSObject.Properties['sharepointIds'] -and $r.sharepointIds -and
                        $r.sharepointIds.PSObject.Properties['listItemUniqueId'] -and
                        $r.sharepointIds.listItemUniqueId) {
                        $r.sharepointIds.listItemUniqueId
                    }
                    elseif ($entityType -eq 'listItem' -and $r.PSObject.Properties['id'] -and $r.id) {
                        $r.id
                    }
                    else { "" }

                    $results.Add([PSCustomObject]@{
                            Source   = "Graph API ($entityType)"
                            ItemType = $entityType
                            Url      = $r.webUrl
                            Owner    = $owner
                            Name     = $itemName
                            UniqueId = $uniqueId
                        })
                }
            }
        }
    }
    catch {
        Write-Host "  ERROR during Graph search: $_" -ForegroundColor Red
        Write-ErrorLog -LogName $log -LogEntryText "Graph search error for $DocumentId : $_"
    }

    # Deduplicate by URL — Graph returns the same file as both driveItem and listItem.
    # Merge UniqueId from the second entry if the first one lacks it.
    $seen = @{}
    $deduped = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($entry in $results) {
        if (-not $seen.ContainsKey($entry.Url)) {
            $seen[$entry.Url] = $deduped.Count
            $deduped.Add($entry)
        }
        elseif ([string]::IsNullOrWhiteSpace($deduped[$seen[$entry.Url]].UniqueId) -and
            -not [string]::IsNullOrWhiteSpace($entry.UniqueId)) {
            $deduped[$seen[$entry.Url]].UniqueId = $entry.UniqueId
        }
    }
    return $deduped
}

# ----------------------------------------------
# Search for a document by name via Microsoft Graph API (tenant-wide)
# ----------------------------------------------
Function Search-DocumentByName {
    param(
        [Parameter(Mandatory)] [string] $DocumentName,
        [Parameter(Mandatory)] [string] $SearchRegion
    )

    $results = [System.Collections.Generic.List[PSObject]]::new()

    try {
        $headers = Get-GraphAuthHeaders
        $searchUrl = "https://graph.microsoft.com/v1.0/search/query"

        # Build a quoted phrase query so partial-word matches are avoided
        $queryString = "`"$DocumentName`""

        foreach ($entityType in @("driveItem", "listItem")) {
            $body = @{
                requests = @(@{
                        entityTypes               = @($entityType)
                        query                     = @{ queryString = $queryString }
                        from                      = 0
                        size                      = 25
                        sharePointOneDriveOptions = @{ includeContent = "sharedContent,privateContent" }
                        region                    = $SearchRegion
                    })
            } | ConvertTo-Json -Depth 5

            $response = Invoke-WithThrottleHandling -ScriptBlock {
                Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Post -Body $body
            } -Operation "Graph name search ($entityType) for '$DocumentName'"

            $hits = $response.value[0].hitsContainers[0].hits
            if ($hits -and $hits.Count -gt 0) {
                foreach ($hit in $hits) {
                    $r = $hit.resource
                    if (-not $r) { continue }

                    $owner = if ($r.createdBy.user.displayName) {
                        if ($r.createdBy.user.email) {
                            "$($r.createdBy.user.displayName) <$($r.createdBy.user.email)>"
                        }
                        else { $r.createdBy.user.displayName }
                    }
                    else { "Unknown" }

                    # driveItem has .name; listItem does not — fall back to the URL leaf
                    $itemName = if ($r.PSObject.Properties['name'] -and $r.name) {
                        $r.name
                    }
                    elseif ($r.webUrl) {
                        try { Split-Path ([System.Uri]::new($r.webUrl).LocalPath) -Leaf } catch { $r.webUrl }
                    }
                    else { "Unknown" }

                    # SharePoint UniqueId is in sharepointIds.listItemUniqueId;
                    # for listItem results, $r.id is also the item GUID
                    $uniqueId = if ($r.PSObject.Properties['sharepointIds'] -and $r.sharepointIds -and
                        $r.sharepointIds.PSObject.Properties['listItemUniqueId'] -and
                        $r.sharepointIds.listItemUniqueId) {
                        $r.sharepointIds.listItemUniqueId
                    }
                    elseif ($entityType -eq 'listItem' -and $r.PSObject.Properties['id'] -and $r.id) {
                        $r.id
                    }
                    else { "" }

                    $results.Add([PSCustomObject]@{
                            Source   = "Graph API ($entityType)"
                            ItemType = $entityType
                            Url      = $r.webUrl
                            Owner    = $owner
                            Name     = $itemName
                            UniqueId = $uniqueId
                        })
                }
            }
        }
    }
    catch {
        Write-Host "  ERROR during Graph name search: $_" -ForegroundColor Red
        Write-ErrorLog -LogName $log -LogEntryText "Graph name search error for '$DocumentName': $_"
    }

    # Deduplicate by URL — Graph returns the same file as both driveItem and listItem.
    # Merge UniqueId from the second entry if the first one lacks it.
    $seen = @{}
    $deduped = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($entry in $results) {
        if (-not $seen.ContainsKey($entry.Url)) {
            $seen[$entry.Url] = $deduped.Count
            $deduped.Add($entry)
        }
        elseif ([string]::IsNullOrWhiteSpace($deduped[$seen[$entry.Url]].UniqueId) -and
            -not [string]::IsNullOrWhiteSpace($entry.UniqueId)) {
            $deduped[$seen[$entry.Url]].UniqueId = $entry.UniqueId
        }
    }
    return $deduped
}
#endregion Search Functions

#region Output
# ============================================================
# Helper: display a result list
# ============================================================
Function Show-Results {
    param([System.Collections.Generic.List[PSObject]] $ResultList)
    foreach ($item in $ResultList) {
        Write-Host "  Source   : $($item.Source)"   -ForegroundColor White
        Write-Host "  Name     : $($item.Name)"     -ForegroundColor White
        Write-Host "  URL      : $($item.Url)"      -ForegroundColor Cyan
        Write-Host "  Owner    : $($item.Owner)"    -ForegroundColor White
        Write-Host "  ItemType : $($item.ItemType)" -ForegroundColor White
        if (-not [string]::IsNullOrWhiteSpace($item.UniqueId)) {
            Write-Host "  UniqueId : $($item.UniqueId)" -ForegroundColor Yellow
        }
        Write-Host ""
    }
}
#endregion Output

#region Main
# ============================================================
# MAIN
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  SharePoint Document Finder" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "How do you want to search?" -ForegroundColor White
Write-Host "  1. By Document Name" -ForegroundColor White
Write-Host "  2. By Document GUID" -ForegroundColor White
Write-Host ""

do {
    $searchChoice = (Read-Host "Enter choice (1 or 2)").Trim()
} while ($searchChoice -notin @('1', '2'))

# --- Collect search term ---
if ($searchChoice -eq '1') {
    do {
        $DocumentName = (Read-Host "Enter the document name to search for").Trim()
    } while ([string]::IsNullOrWhiteSpace($DocumentName))
    Write-InfoLog -LogName $log -LogEntryText "Starting document search by name: '$DocumentName'"
}
else {
    do {
        $DocumentId = (Read-Host "Enter the document GUID to search for").Trim()
        $guidValid = $DocumentId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        if (-not $guidValid) {
            Write-Host "  Invalid GUID format. Expected: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ForegroundColor Yellow
        }
    } while (-not $guidValid)
    Write-InfoLog -LogName $log -LogEntryText "Starting document search by GUID: $DocumentId"
}

# --- Acquire Token ---
Write-Host ""
try {
    AcquireToken
}
catch {
    Write-Host "ERROR: Could not acquire access token: $_" -ForegroundColor Red
    Write-ErrorLog -LogName $log -LogEntryText "Token acquisition failed: $_"
    exit
}

# --- Resolve Search Region ---
if ([string]::IsNullOrWhiteSpace($searchRegion)) {
    $searchRegion = Get-GraphSearchRegion
}
else {
    Write-Host "  Using configured Graph Search region: $searchRegion" -ForegroundColor Cyan
}

# ============================================================
# Branch: Name search
# ============================================================
if ($searchChoice -eq '1') {
    Write-Host ""
    Write-Host "Searching via Microsoft Graph API (tenant-wide)..." -ForegroundColor Cyan
    Write-Host "  Document Name : $DocumentName" -ForegroundColor White
    Write-Host "  Search Region : $searchRegion" -ForegroundColor White
    Write-Host ""

    $graphResults = Search-DocumentByName -DocumentName $DocumentName -SearchRegion $searchRegion

    if ($graphResults -and $graphResults.Count -gt 0) {
        Write-Host "FOUND - $($graphResults.Count) result(s) via Graph API:" -ForegroundColor Green
        Write-Host ""
        Show-Results -ResultList $graphResults
        Write-InfoLog -LogName $log -LogEntryText "Graph API found $($graphResults.Count) result(s) for name '$DocumentName'"
    }
    else {
        Write-Host "NOT FOUND via Graph API." -ForegroundColor Yellow
        Write-Host "The document may not be indexed or the name may not match exactly." -ForegroundColor Yellow
        Write-InfoLog -LogName $log -LogEntryText "Graph API returned no results for name '$DocumentName'"
    }
}
# ============================================================
# Branch: GUID search
# ============================================================
else {
    Write-Host ""
    Write-Host "Searching via Microsoft Graph API (tenant-wide)..." -ForegroundColor Cyan
    Write-Host "  Document GUID : $DocumentId" -ForegroundColor White
    Write-Host "  Search Region : $searchRegion" -ForegroundColor White
    Write-Host ""

    $graphResults = Search-DocumentViaGraphAPI -DocumentId $DocumentId -SearchRegion $searchRegion

    if ($graphResults -and $graphResults.Count -gt 0) {
        Write-Host "FOUND - $($graphResults.Count) result(s) via Graph API:" -ForegroundColor Green
        Write-Host ""
        Show-Results -ResultList $graphResults
        Write-InfoLog -LogName $log -LogEntryText "Graph API found $($graphResults.Count) result(s) for GUID $DocumentId"
    }
    else {
        Write-Host "NOT FOUND via Graph API." -ForegroundColor Yellow
        Write-Host "The document may not be indexed yet, or may have been deleted." -ForegroundColor Yellow
        Write-InfoLog -LogName $log -LogEntryText "Graph API returned no results for GUID $DocumentId"
    }
}

Write-Host ""
if ($debugLogging) {
    Write-Host "Log file: $log" -ForegroundColor DarkGray
}
#endregion Main
