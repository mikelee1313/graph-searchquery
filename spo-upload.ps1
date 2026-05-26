<#
.SYNOPSIS
    Uploads one or more files to a SharePoint document library using the Microsoft Graph API.

.DESCRIPTION
    Authenticates to Microsoft Graph using either a client secret or certificate, then uploads
    a single file or all files in a folder to the specified SharePoint document library.
    Uses simple upload for files under 4MB and resumable upload sessions for larger files.

    Requires an app registration with Sites.ReadWrite.All permission.

.PARAMETER appID
    The application (client) ID of the registered app in Azure AD.

.PARAMETER TenantId
    The tenant ID where the app is registered.

.PARAMETER AuthType
    The authentication type to use. Valid values: 'ClientSecret' or 'Certificate'.

.PARAMETER ClientSecret
    The client secret for the registered app (used when AuthType is 'ClientSecret').

.PARAMETER Thumbprint
    The certificate thumbprint (used when AuthType is 'Certificate').

.PARAMETER SourceFile
    Full path to a single file to upload. Leave blank if using SourceFolder.

.PARAMETER SourceFolder
    Full path to a folder whose contents will be uploaded. Leave blank if using SourceFile.

.PARAMETER UploadSite
    The SharePoint site URL (e.g., "https://contoso.sharepoint.com/sites/Reports").

.PARAMETER UploadLibrary
    The document library name (e.g., "Shared Documents").

.EXAMPLE
    PS> .\spo-upload.ps1

.NOTES
    Author: Mike Lee
    Date: 05/26/26
    Version: 1.0
    - Does not require Microsoft Graph PowerShell SDK for authentication
    - Requires an app with Sites.ReadWrite.All permission
    - Uses simple upload for files < 4MB, resumable session for larger files
#>

#region Configuration
#############################################################
#                  CONFIGURATION SECTION                    #
#############################################################

# Authentication settings - Required
# The app must have Sites.ReadWrite.All permission
$appID = 'abc64618-283f-47ba-a185-50d935d51d57'
$TenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'

# Authentication type: Choose 'ClientSecret' or 'Certificate'
$AuthType = 'Certificate'  # Valid values: 'ClientSecret' or 'Certificate'

# Client Secret authentication (used when $AuthType = 'ClientSecret')
# SECURITY: Prefer environment variable GRAPH_CLIENT_SECRET instead of hardcoding.
# Example (current session): Set-Item Env:\GRAPH_CLIENT_SECRET 'your-secret'
$ClientSecret = $env:GRAPH_CLIENT_SECRET

# Certificate authentication (used when $AuthType = 'Certificate')
$Thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"
$CertStore = "LocalMachine"  # Valid values: 'LocalMachine' or 'CurrentUser'

# Retry / timeout settings for Graph REST calls
$MaxRetries = 10
$InitialBackoffSec = 3
$RequestTimeoutSec = 300

# Source - provide either a single file path or a folder path (not both)
$SourceFile = ""        # Full path to a single file (e.g., "C:\Reports\report.xlsx"). Leave blank to use SourceFolder.
$SourceFolder = ""        # Full path to a folder of files (e.g., "C:\Reports"). Leave blank to use SourceFile.

# SharePoint destination
$UploadSite = "https://m365cpi13246019.sharepoint.com/sites/CopilotReports"  # SharePoint site URL
$UploadLibrary = "Shared Documents"  # Document library name (e.g., "Reports" or "Shared Documents")

#############################################################
#                   END CONFIGURATION                       #
#############################################################
#endregion Configuration

try {

    #region Validation
    if ([string]::IsNullOrWhiteSpace($SourceFile) -and [string]::IsNullOrWhiteSpace($SourceFolder)) {
        throw "You must specify either SourceFile or SourceFolder in the configuration section."
    }
    if (-not [string]::IsNullOrWhiteSpace($SourceFile) -and -not [string]::IsNullOrWhiteSpace($SourceFolder)) {
        throw "Specify either SourceFile or SourceFolder, not both."
    }
    if ([string]::IsNullOrWhiteSpace($UploadSite) -or [string]::IsNullOrWhiteSpace($UploadLibrary)) {
        throw "UploadSite and UploadLibrary must be configured."
    }

    # Build the list of files to upload
    $FilesToUpload = @()
    if (-not [string]::IsNullOrWhiteSpace($SourceFile)) {
        if (-not (Test-Path $SourceFile -PathType Leaf)) {
            throw "SourceFile not found: $SourceFile"
        }
        $FilesToUpload = @($SourceFile)
    }
    else {
        if (-not (Test-Path $SourceFolder -PathType Container)) {
            throw "SourceFolder not found: $SourceFolder"
        }
        $FilesToUpload = Get-ChildItem -Path $SourceFolder -File | Select-Object -ExpandProperty FullName
        if ($FilesToUpload.Count -eq 0) {
            throw "No files found in folder: $SourceFolder"
        }
        Write-Host "Found $($FilesToUpload.Count) file(s) in $SourceFolder" -ForegroundColor Cyan
    }
    #endregion Validation

    #region Initialization
    $global:token = $null
    $global:tokenExpiry = $null
    #endregion Initialization

    #region Helper Functions
    function ConvertTo-Base64Url {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [byte[]]$Bytes
        )
        return [Convert]::ToBase64String($Bytes).TrimEnd('=') -replace '\+', '-' -replace '/', '_'
    }

    function Test-ValidToken {
        [CmdletBinding()]
        param ()

        if ($null -eq $global:tokenExpiry -or (Get-Date) -gt $global:tokenExpiry) {
            Write-Host "Token expired or missing. Refreshing..." -ForegroundColor Yellow
            AcquireToken
        }
    }

    function Get-GraphAuthHeaders {
        [CmdletBinding()]
        param ()

        Test-ValidToken
        return @{ Authorization = "Bearer $global:token" }
    }

    function AcquireToken {
        [CmdletBinding()]
        param ()

        Write-Host "Authenticating to Microsoft Graph using $AuthType..." -ForegroundColor Cyan
        $scope = "https://graph.microsoft.com/.default"
        $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

        if ($AuthType -eq 'ClientSecret') {
            if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
                throw "ClientSecret is empty. Populate ClientSecret in configuration."
            }

            $body = @{
                grant_type    = "client_credentials"
                client_id     = $appID
                client_secret = $ClientSecret
                scope         = $scope
            }

            $response = Invoke-RestMethod -Method POST -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
            $global:token = $response.access_token
            $expiresIn = if ($response.expires_in) { [int]$response.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "Connected via ClientSecret. Token valid until $($global:tokenExpiry)" -ForegroundColor Green
        }
        elseif ($AuthType -eq 'Certificate') {
            $cert = Get-Item -Path "Cert:\$CertStore\My\$Thumbprint" -ErrorAction Stop
            $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
            if (-not $rsa) {
                throw "Unable to access private key for certificate $Thumbprint."
            }

            $now = [System.DateTimeOffset]::UtcNow
            $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
            $nbf = $now.ToUnixTimeSeconds()
            $x5t = ConvertTo-Base64Url -Bytes $cert.GetCertHash()

            $headerJson = @{ alg = "RS256"; typ = "JWT"; x5t = $x5t } | ConvertTo-Json -Compress
            $payloadJson = @{
                aud = $tokenUri
                exp = $exp
                iss = $appID
                jti = [System.Guid]::NewGuid().ToString()
                nbf = $nbf
                sub = $appID
            } | ConvertTo-Json -Compress

            $headerEncoded = ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes($headerJson))
            $payloadEncoded = ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes($payloadJson))
            $unsignedJwt = "$headerEncoded.$payloadEncoded"

            $signatureBytes = $rsa.SignData(
                [System.Text.Encoding]::UTF8.GetBytes($unsignedJwt),
                [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
            )
            $clientAssertion = "$unsignedJwt.$(ConvertTo-Base64Url -Bytes $signatureBytes)"

            $body = @{
                client_id             = $appID
                client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
                client_assertion      = $clientAssertion
                scope                 = $scope
                grant_type            = "client_credentials"
            }

            $response = Invoke-RestMethod -Method POST -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
            $global:token = $response.access_token
            $expiresIn = if ($response.expires_in) { [int]$response.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "Connected via Certificate. Token valid until $($global:tokenExpiry)" -ForegroundColor Green
        }
        else {
            throw "Invalid AuthType '$AuthType'. Valid values are 'ClientSecret' or 'Certificate'."
        }
    }

    function Invoke-GraphRequestWithThrottleHandling {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$Uri,

            [Parameter(Mandatory = $true)]
            [string]$Method,

            [Parameter(Mandatory = $false)]
            [object]$Body = $null,

            [Parameter(Mandatory = $false)]
            [string]$ContentType = "application/json",

            [Parameter(Mandatory = $false)]
            [int]$MaxRetries = $script:MaxRetries,

            [Parameter(Mandatory = $false)]
            [int]$InitialBackoffSeconds = $script:InitialBackoffSec,

            [Parameter(Mandatory = $false)]
            [int]$TimeoutSeconds = $script:RequestTimeoutSec
        )

        $retryCount = 0
        $backoffSec = $InitialBackoffSeconds

        while ($retryCount -le $MaxRetries) {
            try {
                $headers = Get-GraphAuthHeaders
                $invokeParams = @{
                    Uri         = $Uri
                    Method      = $Method
                    Headers     = $headers
                    ContentType = $ContentType
                    TimeoutSec  = $TimeoutSeconds
                    ErrorAction = "Stop"
                }
                if ($null -ne $Body) {
                    $invokeParams["Body"] = $Body
                }
                return Invoke-RestMethod @invokeParams
            }
            catch {
                $statusCode = $null
                $retryAfterSec = $null

                if ($_.Exception.Response) {
                    try {
                        $statusCode = [int]$_.Exception.Response.StatusCode
                    }
                    catch {
                        $statusCode = $_.Exception.Response.StatusCode.value__
                    }

                    if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers.Contains("Retry-After")) {
                        $retryAfterRaw = $_.Exception.Response.Headers.GetValues("Retry-After") | Select-Object -First 1
                        if ($retryAfterRaw -match '^\d+$') {
                            $retryAfterSec = [int]$retryAfterRaw
                        }
                        else {
                            $retryAfterDate = [datetime]::MinValue
                            if ([datetime]::TryParse($retryAfterRaw, [ref]$retryAfterDate)) {
                                $delay = [math]::Ceiling(($retryAfterDate.ToUniversalTime() - (Get-Date).ToUniversalTime()).TotalSeconds)
                                if ($delay -gt 0) {
                                    $retryAfterSec = [int]$delay
                                }
                            }
                        }
                    }
                }

                $isRetryable = $statusCode -in @(429, 502, 503, 504)
                if (-not $isRetryable -or $retryCount -ge $MaxRetries) {
                    throw
                }

                $retryCount++
                $waitSec = if ($retryAfterSec) { $retryAfterSec } else { $backoffSec }
                Write-Host "Graph request throttled/transient error ($statusCode). Retrying in $waitSec sec (attempt $retryCount/$MaxRetries)..." -ForegroundColor Yellow
                Start-Sleep -Seconds $waitSec
                $backoffSec = [Math]::Min($backoffSec * 2, 300)
            }
        }
    }

    function Invoke-ChunkUploadWithRetry {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$UploadUrl,

            [Parameter(Mandatory = $true)]
            [byte[]]$ChunkData,

            [Parameter(Mandatory = $true)]
            [string]$ContentRange,

            [Parameter(Mandatory = $true)]
            [int]$ContentLength,

            [Parameter(Mandatory = $false)]
            [int]$MaxRetries = $script:MaxRetries,

            [Parameter(Mandatory = $false)]
            [int]$InitialBackoffSeconds = $script:InitialBackoffSec,

            [Parameter(Mandatory = $false)]
            [int]$TimeoutSeconds = $script:RequestTimeoutSec
        )

        $retryCount = 0
        $backoffSec = $InitialBackoffSeconds

        while ($retryCount -le $MaxRetries) {
            try {
                $headers = @{
                    "Content-Length" = $ContentLength.ToString()
                    "Content-Range"  = $ContentRange
                }

                return Invoke-RestMethod -Uri $UploadUrl -Method PUT -Body $ChunkData -Headers $headers -ContentType "application/octet-stream" -TimeoutSec $TimeoutSeconds -ErrorAction Stop
            }
            catch {
                $statusCode = $null
                $retryAfterSec = $null

                if ($_.Exception.Response) {
                    try {
                        $statusCode = [int]$_.Exception.Response.StatusCode
                    }
                    catch {
                        $statusCode = $_.Exception.Response.StatusCode.value__
                    }

                    if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers.Contains("Retry-After")) {
                        $retryAfterRaw = $_.Exception.Response.Headers.GetValues("Retry-After") | Select-Object -First 1
                        if ($retryAfterRaw -match '^\d+$') {
                            $retryAfterSec = [int]$retryAfterRaw
                        }
                        else {
                            $retryAfterDate = [datetime]::MinValue
                            if ([datetime]::TryParse($retryAfterRaw, [ref]$retryAfterDate)) {
                                $delay = [math]::Ceiling(($retryAfterDate.ToUniversalTime() - (Get-Date).ToUniversalTime()).TotalSeconds)
                                if ($delay -gt 0) {
                                    $retryAfterSec = [int]$delay
                                }
                            }
                        }
                    }
                }

                $isRetryable = $statusCode -in @(429, 500, 502, 503, 504)
                if (-not $isRetryable -or $retryCount -ge $MaxRetries) {
                    throw
                }

                $retryCount++
                $waitSec = if ($retryAfterSec) { $retryAfterSec } else { $backoffSec }
                Write-Host "Chunk upload transient error ($statusCode). Retrying in $waitSec sec (attempt $retryCount/$MaxRetries)..." -ForegroundColor Yellow
                Start-Sleep -Seconds $waitSec
                $backoffSec = [Math]::Min($backoffSec * 2, 300)
            }
        }
    }
    #endregion Helper Functions

    #region Authentication
    try {
        AcquireToken
    }
    catch {
        Write-Host "Authentication failed:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        throw
    }
    #endregion Authentication

    #region SharePoint Upload
    Write-Host ""
    Write-Host "=========================================================="
    Write-Host "UPLOADING TO SHAREPOINT"
    Write-Host "=========================================================="
    Write-Host "Site: $UploadSite" -ForegroundColor Gray
    Write-Host "Library: $UploadLibrary" -ForegroundColor Gray

    try {
        # Resolve site ID
        $SiteUrl = $UploadSite -replace "https://", ""
        $SiteParts = $SiteUrl -split "/"
        $HostName = $SiteParts[0]
        $SitePath = if ($SiteParts.Count -gt 1) { $SiteParts[1..($SiteParts.Count - 1)] -join "/" } else { "" }

        $SiteUri = "https://graph.microsoft.com/v1.0/sites/$HostName`:/$SitePath"
        $Site = Invoke-GraphRequestWithThrottleHandling -Uri $SiteUri -Method GET
        $SiteId = $Site.id
        Write-Host "Site ID: $SiteId" -ForegroundColor Gray

        # Resolve drive ID
        $DriveUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives?`$filter=name eq '$UploadLibrary'"
        $DriveResponse = Invoke-GraphRequestWithThrottleHandling -Uri $DriveUri -Method GET

        if ($DriveResponse.value.Count -eq 0) {
            Write-Host "Document library '$UploadLibrary' not found. Falling back to default drive..." -ForegroundColor Yellow
            $Drive = Invoke-GraphRequestWithThrottleHandling -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drive" -Method GET
            $DriveId = $Drive.id
        }
        else {
            $DriveId = $DriveResponse.value[0].id
        }
        Write-Host "Drive ID: $DriveId" -ForegroundColor Gray
    }
    catch {
        Write-Host "Error resolving SharePoint site or library:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        throw
    }

    # Upload each file
    $SuccessCount = 0
    $FailCount = 0

    foreach ($FileToUpload in $FilesToUpload) {
        $FileName = Split-Path $FileToUpload -Leaf
        $EncodedFileName = [uri]::EscapeDataString($FileName)
        $FileInfo = Get-Item $FileToUpload
        $FileSizeBytes = $FileInfo.Length
        $FileSizeMB = [math]::Round($FileSizeBytes / 1MB, 2)

        Write-Host ""
        Write-Host "Uploading: $FileName ($FileSizeMB MB)..." -ForegroundColor Cyan

        try {
            if ($FileSizeBytes -lt 4MB) {
                # Simple upload for files under 4MB
                Write-Host "  Using simple upload method (file < 4MB)..." -ForegroundColor Gray
                $FileContent = [System.IO.File]::ReadAllBytes($FileToUpload)
                $UploadUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$EncodedFileName`:/content"
                $UploadResult = Invoke-GraphRequestWithThrottleHandling -Uri $UploadUri -Method PUT -Body $FileContent -ContentType "application/octet-stream"
                Write-Host "  Uploaded successfully. URL: $($UploadResult.webUrl)" -ForegroundColor Green
            }
            else {
                # Resumable upload session for files 4MB or larger
                Write-Host "  Using resumable upload session (file >= 4MB)..." -ForegroundColor Gray
                $CreateSessionUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$EncodedFileName`:/createUploadSession"
                $SessionBody = @{
                    item = @{
                        "@microsoft.graph.conflictBehavior" = "replace"
                    }
                } | ConvertTo-Json

                $UploadSession = Invoke-GraphRequestWithThrottleHandling -Uri $CreateSessionUri -Method POST -Body $SessionBody -ContentType "application/json"
                $UploadUrl = $UploadSession.uploadUrl

                $ChunkSize = 10MB
                $FileStream = [System.IO.File]::OpenRead($FileToUpload)
                $Buffer = New-Object byte[] $ChunkSize
                $BytesUploaded = 0
                $ChunkNumber = 0
                $UploadResult = $null

                try {
                    while (($BytesRead = $FileStream.Read($Buffer, 0, $ChunkSize)) -gt 0) {
                        $ChunkNumber++
                        $StartByte = $BytesUploaded
                        $EndByte = $BytesUploaded + $BytesRead - 1

                        if ($BytesRead -lt $ChunkSize) {
                            $ChunkData = New-Object byte[] $BytesRead
                            [Array]::Copy($Buffer, $ChunkData, $BytesRead)
                        }
                        else {
                            $ChunkData = $Buffer
                        }

                        $ContentRange = "bytes $StartByte-$EndByte/$FileSizeBytes"
                        $PercentComplete = [math]::Round(($EndByte / $FileSizeBytes) * 100, 1)
                        Write-Host "  Chunk $ChunkNumber - $ContentRange ($PercentComplete%)" -ForegroundColor Gray

                        $UploadResult = Invoke-ChunkUploadWithRetry -UploadUrl $UploadUrl -ChunkData $ChunkData -ContentRange $ContentRange -ContentLength $BytesRead
                        $BytesUploaded += $BytesRead
                    }
                    Write-Host "  Uploaded successfully. URL: $($UploadResult.webUrl)" -ForegroundColor Green
                }
                finally {
                    $FileStream.Close()
                    $FileStream.Dispose()
                }
            }
            $SuccessCount++
        }
        catch {
            Write-Host "  ERROR uploading $FileName`: $($_.Exception.Message)" -ForegroundColor Red
            $FailCount++
        }
    }

    Write-Host ""
    Write-Host "=========================================================="
    Write-Host "Upload complete: $SuccessCount succeeded, $FailCount failed." -ForegroundColor $(if ($FailCount -eq 0) { 'Green' } else { 'Yellow' })
    Write-Host "=========================================================="
    #endregion SharePoint Upload

}
catch {
    Write-Host "FATAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
