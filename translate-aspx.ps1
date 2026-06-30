<#
.SYNOPSIS
    Translates SharePoint ASPX pages and creates translated pages in the same Site Pages library.

.DESCRIPTION
    This script reads modern SharePoint ASPX pages with Microsoft Graph, translates page titles,
    text web parts, title-area text, and common web part text properties with Azure AI Translator,
    then creates a new SharePoint page in the same Site Pages library with Microsoft Graph.

    Microsoft Graph is used for SharePoint page read/create operations. Azure AI Translator is used
    for language translation because Microsoft Graph does not provide a general page translation API.

.PARAMETER tenantId
    The Azure AD tenant ID for authentication.

.PARAMETER clientId
    The client ID of the Azure AD application.

.PARAMETER Thumbprint
    The certificate thumbprint for authentication.

.PARAMETER AppAuthMode
    Application authentication mode for Microsoft Graph and Entra-based Translator auth.
    Use "Certificate" (default) to sign with a certificate private key.
    Use "ClientSecret" to authenticate with an app secret and avoid private key prompts.

.PARAMETER ClientSecret
    The Azure AD application client secret. Defaults to AZURE_CLIENT_SECRET environment variable.
    Required when AppAuthMode is "ClientSecret".

.PARAMETER CertStore
    Certificate store location: "LocalMachine" or "CurrentUser" (default: "LocalMachine").

.PARAMETER siteUrl
    The SharePoint site URL that contains the Site Pages library.

.PARAMETER PageName
    Optional input page reference. Accepts the page file name, page title, or page URL.
    If omitted, all ASPX pages returned by Graph are translated.

.PARAMETER TargetLanguage
    Translator target language code. Defaults to "fr" for French.

.PARAMETER SourceLanguage
    Optional Translator source language code. If omitted, Translator auto-detects the source language.

.PARAMETER OutputNameSuffix
    Suffix added to the new page file name. Defaults to "-fr".

.PARAMETER AzureTranslatorKey
    Azure AI Translator key. Defaults to the AZURE_TRANSLATOR_KEY environment variable.

.PARAMETER AzureTranslatorEndpoint
    Azure AI Translator endpoint. Defaults to https://api.cognitive.microsofttranslator.com.

.PARAMETER AzureTranslatorRegion
    Azure AI Translator region. Defaults to the AZURE_TRANSLATOR_REGION environment variable.

.PARAMETER TranslatorTenantId
    Optional tenant ID used only for Entra token acquisition to Azure AI Translator.
    Use this when Microsoft Graph and Translator are in different tenants.
    If omitted, tenantId is used for both services.

.PARAMETER TranslatorAuthMode
    Translator authentication mode. Use "Entra" when API key based authentication is disabled.
    Use "Key" to authenticate with AzureTranslatorKey.

.PARAMETER Draft
    Leaves each translated page as a draft after creation. By default, translated pages are published so they appear in Site Pages.

.EXAMPLE
    .\translate-aspx.ps1 -PageName "Home.aspx" -TargetLanguage fr
    Creates and publishes a French page for Home.aspx in the same Site Pages library.

.NOTES
    Author: Mike Lee
    Created: 06/25/2026
    Updated: 6/29/2026 - Added ClientSecret support 
    Updated: 6/30/2026 - Added support for translations from subfolders (Via REST)

    Required setup:
    1. Grant the app registration Microsoft Graph application permission Sites.ReadWrite.All and admin consent.
    2. For Entra Translator auth, assign the app registration/service principal Azure RBAC on the Azure AI Services resource:
       Azure portal > Azure AI Services resource > Access control (IAM) > Add role assignment > Cognitive Services Contributor.
       If your tenant allows it, Cognitive Services User may also work; use Cognitive Services Contributor if Translator returns
       "PermissionDenied: Principal does not have access to API/Operation."
       In the member picker, search for and select the Enterprise Application display name for the same app/client ID used by
       this script. The picker may not find the app when searching by client ID directly.
       Wait a few minutes for RBAC propagation after assigning the role.
    3. If using key auth instead of Entra auth, enable "Allow API key based authentication" on the Azure AI Services resource
       and run with -TranslatorAuthMode Key plus a valid AzureTranslatorKey.

    Required Graph application permission: Sites.ReadWrite.All
    Required Azure service: Azure AI Translator
    Required Azure RBAC for Entra Translator auth: assign the app registration/service principal Cognitive Services Contributor on the Azure AI Services resource.
#>

[CmdletBinding()]
param(
    [string]$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3',
    [string]$clientId = 'abc64618-283f-47ba-a185-50d935d51d57',
    [ValidateSet('Certificate', 'ClientSecret')]
    [string]$AppAuthMode = 'Certificate',
    [string]$Thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9",
    [string]$ClientSecret = "",
    [ValidateSet('LocalMachine', 'CurrentUser')]
    [string]$CertStore = 'LocalMachine',
    [string]$siteUrl = "https://m365cpi13246019.sharepoint.com/sites/SalesandMarketing",
    [string]$PageName = "Testpageinfolder-eng2.aspx",
    [string]$TargetLanguage = "fr",
    [string]$SourceLanguage = "",
    [string]$OutputNameSuffix = "-fr",
    [string]$AzureTranslatorKey = $env:AZURE_TRANSLATOR_KEY,
    [string]$AzureTranslatorEndpoint = 'https://m365cpi13246019azureservices.cognitiveservices.azure.com/',
    [string]$AzureTranslatorRegion = 'eastus',
    [string]$TranslatorTenantId = '',
    [ValidateSet('Entra', 'Key')]
    [string]$TranslatorAuthMode = 'Entra',
    [switch]$Draft
)

#region Initialization

$date = Get-Date -Format "yyyyMMddHHmmss"
$logFolder = Join-Path -Path $env:TEMP -ChildPath "ASPX_Translation_Logs"

if (-not (Test-Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
}

$logFile = Join-Path -Path $logFolder -ChildPath "translation_log_$date.log"
$csvLog = Join-Path -Path $logFolder -ChildPath "translation_summary_$date.csv"

$global:token = ""
$global:sharePointToken = ""
$global:cognitiveServicesToken = ""
$global:translationResults = @()
$script:translationCache = @{}
$script:sitePagesDriveIds = @{}

#endregion

#region Logging Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'INFO' { Write-Host $logEntry -ForegroundColor Cyan }
        'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
        'ERROR' { Write-Host $logEntry -ForegroundColor Red }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
    }

    Add-Content -Path $logFile -Value $logEntry
}

function Add-TranslationResult {
    param(
        [string]$SourcePage,
        [string]$SourceWebUrl,
        [string]$TargetPage,
        [string]$TargetWebUrl = "",
        [string]$Status,
        [string]$ErrorMessage = "",
        [int]$WebPartsProcessed = 0,
        [int]$TextSegmentsTranslated = 0,
        [datetime]$StartTime,
        [datetime]$EndTime
    )

    $result = [PSCustomObject]@{
        Timestamp              = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        SourcePage             = $SourcePage
        SourceWebUrl           = $SourceWebUrl
        TargetPage             = $TargetPage
        TargetWebUrl           = $TargetWebUrl
        TargetLanguage         = $TargetLanguage
        Status                 = $Status
        ErrorMessage           = $ErrorMessage
        WebPartsProcessed      = $WebPartsProcessed
        TextSegmentsTranslated = $TextSegmentsTranslated
        ProcessingTime         = ($EndTime - $StartTime).TotalSeconds
        StartTime              = $StartTime.ToString("yyyy-MM-dd HH:mm:ss")
        EndTime                = $EndTime.ToString("yyyy-MM-dd HH:mm:ss")
    }

    $global:translationResults += $result
}

#endregion

#region Authentication

function Get-CertificateAccessToken {
    param(
        [string]$Scope,
        [string]$ServiceName,
        [string]$AuthorityTenantId = $tenantId
    )

    Write-Log "Requesting $ServiceName token using certificate authentication..."
    $certPath = "Cert:\$CertStore\My\$Thumbprint"
    if (-not (Test-Path $certPath)) {
        Write-Log "Certificate not found at $certPath" -Level ERROR
        exit 1
    }

    $cert = Get-Item $certPath
    Write-Log "Certificate found in $CertStore\My store" -Level SUCCESS

    $tokenUrl = "https://login.microsoftonline.com/$AuthorityTenantId/oauth2/v2.0/token"

    $jwtHeader = @{
        alg = "RS256"
        typ = "JWT"
        x5t = [Convert]::ToBase64String($cert.GetCertHash()) -replace '\+', '-' -replace '/', '_' -replace '='
    }

    $jwtPayload = @{
        aud = $tokenUrl
        exp = ([DateTimeOffset]::UtcNow.AddMinutes(5).ToUnixTimeSeconds())
        iss = $clientId
        jti = [guid]::NewGuid().ToString()
        nbf = ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
        sub = $clientId
    }

    $headerJson = $jwtHeader | ConvertTo-Json -Compress
    $payloadJson = $jwtPayload | ConvertTo-Json -Compress

    $headerBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($headerJson)) -replace '\+', '-' -replace '/', '_' -replace '='
    $payloadBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($payloadJson)) -replace '\+', '-' -replace '/', '_' -replace '='

    $dataToSign = "$headerBase64.$payloadBase64"
    $dataToSignBytes = [Text.Encoding]::UTF8.GetBytes($dataToSign)

    $signature = $cert.PrivateKey.SignData($dataToSignBytes, [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $signatureBase64 = [Convert]::ToBase64String($signature) -replace '\+', '-' -replace '/', '_' -replace '='

    $clientAssertion = "$dataToSign.$signatureBase64"

    $body = @{
        client_id             = $clientId
        client_assertion      = $clientAssertion
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        scope                 = $Scope
        grant_type            = "client_credentials"
    }

    $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
    Write-Log "Successfully acquired $ServiceName token" -Level SUCCESS

    return $response.access_token
}

function Get-ClientSecretAccessToken {
    param(
        [string]$Scope,
        [string]$ServiceName,
        [string]$AuthorityTenantId = $tenantId
    )

    Write-Log "Requesting $ServiceName token using client secret authentication..."

    if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
        throw "ClientSecret is required when AppAuthMode is ClientSecret. Pass -ClientSecret or set AZURE_CLIENT_SECRET."
    }

    $tokenUrl = "https://login.microsoftonline.com/$AuthorityTenantId/oauth2/v2.0/token"

    $body = @{
        client_id     = $clientId
        client_secret = $ClientSecret
        scope         = $Scope
        grant_type    = "client_credentials"
    }

    $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
    Write-Log "Successfully acquired $ServiceName token" -Level SUCCESS

    return $response.access_token
}

function Get-AppAccessToken {
    param(
        [string]$Scope,
        [string]$ServiceName,
        [string]$AuthorityTenantId = $tenantId
    )

    if ($AppAuthMode -eq 'ClientSecret') {
        return Get-ClientSecretAccessToken -Scope $Scope -ServiceName $ServiceName -AuthorityTenantId $AuthorityTenantId
    }

    return Get-CertificateAccessToken -Scope $Scope -ServiceName $ServiceName -AuthorityTenantId $AuthorityTenantId
}

function Get-EffectiveTranslatorTenantId {
    if ([string]::IsNullOrWhiteSpace($TranslatorTenantId)) {
        return $tenantId
    }

    return $TranslatorTenantId.Trim()
}

function Get-GraphToken {
    $global:token = Get-AppAccessToken -Scope "https://graph.microsoft.com/.default" -ServiceName "Microsoft Graph" -AuthorityTenantId $tenantId
}

function Get-SharePointToken {
    $siteUri = [System.Uri]$siteUrl
    $global:sharePointToken = Get-AppAccessToken -Scope "https://$($siteUri.Host)/.default" -ServiceName "SharePoint" -AuthorityTenantId $tenantId
}

function Get-CognitiveServicesToken {
    $global:cognitiveServicesToken = Get-AppAccessToken -Scope "https://cognitiveservices.azure.com/.default" -ServiceName "Azure AI Services" -AuthorityTenantId (Get-EffectiveTranslatorTenantId)
}

function New-GraphHeaders {
    param(
        [switch]$MetadataNone
    )

    $headers = @{
        "Authorization" = "Bearer $global:token"
        "Content-Type"  = "application/json"
    }

    if ($MetadataNone) {
        $headers["Accept"] = "application/json;odata.metadata=none"
    }

    return $headers
}

function New-SharePointHeaders {
    return @{
        "Authorization" = "Bearer $global:sharePointToken"
        "Accept"        = "application/json;odata=nometadata"
        "Content-Type"  = "application/json;odata=nometadata"
    }
}

#endregion

#region Request Helpers

function Invoke-GraphRequestWithRetry {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Method = "GET",
        [object]$Body = $null,
        [string]$ContentType = "application/json",
        [int]$MaxRetries = 5
    )

    $retryCount = 0
    $baseDelay = 1

    while ($retryCount -le $MaxRetries) {
        try {
            $params = @{
                Uri         = $Uri
                Headers     = $Headers
                Method      = $Method
                ContentType = $ContentType
            }

            if ($null -ne $Body) {
                $params['Body'] = $Body
            }

            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) {
                $retryAfter = [Math]::Pow(2, $retryCount) * $baseDelay

                if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers['Retry-After']) {
                    $retryAfter = [int]$_.Exception.Response.Headers['Retry-After']
                }

                if ($retryCount -lt $MaxRetries) {
                    Write-Log "Graph request returned $statusCode. Retrying in $retryAfter seconds ($($retryCount + 1)/$MaxRetries)..." -Level WARNING
                    Start-Sleep -Seconds $retryAfter
                    $retryCount++
                    continue
                }
            }

            throw
        }
    }
}

function Invoke-SharePointRequestWithRetry {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Method = "GET",
        [object]$Body = $null,
        [string]$ContentType = "application/json;odata=nometadata",
        [int]$MaxRetries = 5
    )

    $retryCount = 0
    $baseDelay = 1

    while ($retryCount -le $MaxRetries) {
        try {
            $params = @{
                Uri         = $Uri
                Headers     = $Headers
                Method      = $Method
                ContentType = $ContentType
            }

            if ($null -ne $Body) {
                $params['Body'] = $Body
            }

            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) {
                $retryAfter = [Math]::Pow(2, $retryCount) * $baseDelay

                if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers['Retry-After']) {
                    $retryAfter = [int]$_.Exception.Response.Headers['Retry-After']
                }

                if ($retryCount -lt $MaxRetries) {
                    Write-Log "SharePoint REST request returned $statusCode. Retrying in $retryAfter seconds ($($retryCount + 1)/$MaxRetries)..." -Level WARNING
                    Start-Sleep -Seconds $retryAfter
                    $retryCount++
                    continue
                }
            }

            throw
        }
    }
}

function Invoke-TranslatorRequestWithRetry {
    param(
        [string]$Uri,
        [string]$Body,
        [int]$MaxRetries = 5
    )

    $translatorRegion = Get-NormalizedTranslatorRegion

    $headers = @{
        "Content-Type"    = "application/json; charset=utf-8"
        "X-ClientTraceId" = [guid]::NewGuid().ToString()
    }

    if ($TranslatorAuthMode -eq 'Entra') {
        if ([string]::IsNullOrWhiteSpace($global:cognitiveServicesToken)) {
            throw "Azure AI Services Entra token is missing. Call Get-CognitiveServicesToken before Translator requests."
        }

        $headers["Authorization"] = "Bearer $global:cognitiveServicesToken"
    }
    else {
        $headers["Ocp-Apim-Subscription-Key"] = $AzureTranslatorKey.Trim()
    }

    if ($translatorRegion) {
        $headers["Ocp-Apim-Subscription-Region"] = $translatorRegion
    }

    $retryCount = 0
    $baseDelay = 1

    while ($retryCount -le $MaxRetries) {
        try {
            return Invoke-RestMethod -Uri $Uri -Method POST -Headers $headers -Body $Body -ContentType "application/json; charset=utf-8"
        }
        catch {
            $statusCode = $null
            $responseBody = ""
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $responseBody = $_.ErrorDetails.Message
            }
            elseif ($_.Exception.Response -and $_.Exception.Response.GetResponseStream()) {
                $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $responseBody = $reader.ReadToEnd()
                $reader.Dispose()
            }

            if ($statusCode -eq 401 -or $statusCode -eq 403) {
                $regionGuidance = if ($translatorRegion) {
                    "Verify AZURE_TRANSLATOR_REGION matches the Azure Translator resource region ('$translatorRegion')."
                }
                else {
                    "If this is a regional Azure Translator resource, set -AzureTranslatorRegion or AZURE_TRANSLATOR_REGION. Only global Translator resources omit the region header."
                }

                if ($TranslatorAuthMode -eq 'Entra') {
                    throw "Azure AI Translator authorization failed ($statusCode) for endpoint '$Uri' using Entra auth. Response body: $responseBody Verify the app registration/service principal has Azure RBAC such as Cognitive Services Contributor on the Azure AI Services resource. If it only has Cognitive Services User and Translator returns PermissionDenied, add Cognitive Services Contributor. $regionGuidance"
                }

                throw "Azure AI Translator authorization failed ($statusCode) for endpoint '$Uri' using key auth. Response body: $responseBody Verify -AzureTranslatorKey/AZURE_TRANSLATOR_KEY is valid and belongs to the resource. $regionGuidance"
            }

            if ($statusCode -eq 400) {
                if ($TranslatorAuthMode -eq 'Entra' -and $responseBody -match 'Token tenant .* does not match resource tenant|Tenant provided in token does not match resource token') {
                    $effectiveTranslatorTenantId = Get-EffectiveTranslatorTenantId
                    throw "Azure AI Translator tenant mismatch (400). Response body: $responseBody The token tenant does not match the Translator resource tenant. If Graph and Translator are in different tenants, set -TranslatorTenantId to the Translator resource tenant (current effective Translator tenant: '$effectiveTranslatorTenantId'). If using API key auth, run with -TranslatorAuthMode Key so no Entra token is sent."
                }

                throw "Azure AI Translator request failed (400) for endpoint '$Uri'. Response body: $responseBody"
            }

            if ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) {
                $retryAfter = [Math]::Pow(2, $retryCount) * $baseDelay

                if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers['Retry-After']) {
                    $retryAfter = [int]$_.Exception.Response.Headers['Retry-After']
                }

                if ($retryCount -lt $MaxRetries) {
                    Write-Log "Translator request returned $statusCode. Retrying in $retryAfter seconds ($($retryCount + 1)/$MaxRetries)..." -Level WARNING
                    Start-Sleep -Seconds $retryAfter
                    $retryCount++
                    continue
                }
            }

            throw
        }
    }
}

function Get-GraphPagedResults {
    param(
        [string]$Uri,
        [hashtable]$Headers
    )

    $results = @()
    $nextLink = $Uri

    while ($nextLink) {
        $response = Invoke-GraphRequestWithRetry -Uri $nextLink -Headers $Headers -Method GET
        if ($response.value) {
            $results += $response.value
        }
        $nextLink = $response.'@odata.nextLink'
    }

    return $results
}

#endregion

#region Translation Helpers

function Get-NormalizedTranslatorRegion {
    if ([string]::IsNullOrWhiteSpace($AzureTranslatorRegion)) {
        return ""
    }

    return (($AzureTranslatorRegion.Trim() -replace '\s', '').ToLowerInvariant())
}

function Get-TranslatorTranslateUri {
    param(
        [string]$TextType
    )

    $endpoint = $AzureTranslatorEndpoint.TrimEnd('/')
    $endpointUri = [System.Uri]$endpoint
    $baseUri = $endpoint

    if ($endpointUri.Host -like "*.cognitiveservices.azure.com" -and $endpointUri.AbsolutePath -notmatch "/translator/text/v3\.0/?$") {
        $baseUri = "$endpoint/translator/text/v3.0"
    }

    $encodedTo = [System.Web.HttpUtility]::UrlEncode($TargetLanguage.Trim())
    $uri = "$baseUri/translate?api-version=3.0&to=$encodedTo&textType=$TextType"

    if (-not [string]::IsNullOrWhiteSpace($SourceLanguage)) {
        $encodedFrom = [System.Web.HttpUtility]::UrlEncode($SourceLanguage.Trim())
        $uri += "&from=$encodedFrom"
    }

    return $uri
}

function Assert-TranslatorConfigured {
    if ($AppAuthMode -eq 'ClientSecret') {
        if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
            Write-Log "ClientSecret is required when AppAuthMode is ClientSecret. Pass -ClientSecret or set AZURE_CLIENT_SECRET." -Level ERROR
            exit 1
        }
    }
    else {
        $certPath = "Cert:\$CertStore\My\$Thumbprint"
        if (-not (Test-Path $certPath)) {
            Write-Log "Certificate not found at $certPath" -Level ERROR
            exit 1
        }
    }

    if ([string]::IsNullOrWhiteSpace($TargetLanguage)) {
        Write-Log "TargetLanguage is required." -Level ERROR
        exit 1
    }

    if ($TranslatorAuthMode -eq 'Key') {
        if ([string]::IsNullOrWhiteSpace($AzureTranslatorKey)) {
            Write-Log "AzureTranslatorKey is required when TranslatorAuthMode is Key. Pass -AzureTranslatorKey or set AZURE_TRANSLATOR_KEY." -Level ERROR
            exit 1
        }

        if ($AzureTranslatorKey -match '[<>]' -or $AzureTranslatorKey -match 'your-|<key>|translator-key') {
            Write-Log "AzureTranslatorKey still appears to contain a placeholder value." -Level ERROR
            exit 1
        }

        $trimmedTranslatorKey = $AzureTranslatorKey.Trim()
        if ($trimmedTranslatorKey -match '\.\.\.|…') {
            Write-Log "AzureTranslatorKey appears to contain the Azure portal's truncated display value ('...' or '…'). Use the copy button next to Primary Key/Key 1 instead of selecting the visible text." -Level ERROR
            exit 1
        }
    }

    $translatorRegion = Get-NormalizedTranslatorRegion

    Write-Log "Translator endpoint: $($AzureTranslatorEndpoint.TrimEnd('/'))"
    Write-Log "Translator translate URL pattern: $(Get-TranslatorTranslateUri -TextType plain)"
    Write-Log "Translator region header: $(if ($translatorRegion) { $translatorRegion } else { '<not set>' })"
    Write-Log "Translator auth mode: $TranslatorAuthMode"
    if ($TranslatorAuthMode -eq 'Entra') {
        Write-Log "Translator token tenant: $(Get-EffectiveTranslatorTenantId)"
    }
    Write-Log "Translator key configured: $(if ($AzureTranslatorKey) { 'yes' } else { 'no' })"
}

function Test-TranslatorCredentials {
    Write-Log "Validating Azure AI Translator credentials..."

    try {
        $testTranslation = Translate-Text -Text "Translator credential validation" -TextType plain
        if ([string]::IsNullOrWhiteSpace($testTranslation)) {
            throw "Translator returned an empty plain-text validation response."
        }

        $testHtmlTranslation = Translate-Text -Text "<p>Translator HTML credential validation</p>" -TextType html
        if ([string]::IsNullOrWhiteSpace($testHtmlTranslation)) {
            throw "Translator returned an empty HTML validation response."
        }

        Write-Log "Azure AI Translator credentials validated" -Level SUCCESS
    }
    catch {
        Write-Log $_.Exception.Message -Level ERROR
        if ($TranslatorAuthMode -eq 'Entra') {
            Write-Log "For Entra auth, 401/403 PermissionDenied means the app's Azure RBAC is missing, assigned to the wrong Enterprise Application, assigned at the wrong scope, or has not propagated yet. Add Cognitive Services Contributor if Cognitive Services User is not sufficient." -Level ERROR
        }
        else {
            Write-Log "Set credentials with: `$env:AZURE_TRANSLATOR_KEY = '<key>'; `$env:AZURE_TRANSLATOR_REGION = '<region>'" -Level ERROR
            Write-Log "You can also pass -AzureTranslatorKey and -AzureTranslatorRegion directly when running the script." -Level ERROR
        }
        exit 1
    }
}

function Test-TranslatableText {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    if ($Text -match '^\s*(https?://|/sites/|/_layouts/|#|\{|\[)') {
        return $false
    }

    if ($Text -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        return $false
    }

    if ($Text -notmatch '[A-Za-z]') {
        return $false
    }

    return $true
}

function Test-TranslatableProperty {
    param(
        [string]$PropertyName
    )

    if ([string]::IsNullOrWhiteSpace($PropertyName)) {
        return $false
    }

    $translatableNames = @(
        'title',
        'description',
        'altText',
        'overlayText',
        'captionText',
        'buttonText',
        'label',
        'ariaLabel',
        'text',
        'textAboveTitle',
        'placeholderText',
        'callToActionText',
        'searchablePlainText'
    )

    return $translatableNames -contains $PropertyName
}

function Translate-Text {
    param(
        [string]$Text,
        [ValidateSet('plain', 'html')]
        [string]$TextType = 'plain',
        [string]$Context = ""
    )

    if (-not (Test-TranslatableText -Text $Text)) {
        return $Text
    }

    $cacheKey = "$TextType|$SourceLanguage|$TargetLanguage|$Text"
    if ($script:translationCache.ContainsKey($cacheKey)) {
        return $script:translationCache[$cacheKey]
    }

    $uri = Get-TranslatorTranslateUri -TextType $TextType
    $body = ConvertTo-Json -InputObject @(@{ Text = $Text }) -Depth 5 -Compress
    try {
        $response = Invoke-TranslatorRequestWithRetry -Uri $uri -Body $body
    }
    catch {
        $contextMessage = if ($Context) { " Context: $Context." } else { "" }
        throw "$($_.Exception.Message) TextType: $TextType. Text length: $($Text.Length).$contextMessage"
    }

    if (-not $response -or -not $response[0].translations -or -not $response[0].translations[0].text) {
        throw "Translator returned an unexpected response for text segment."
    }

    $translated = $response[0].translations[0].text
    $script:translationCache[$cacheKey] = $translated
    return $translated
}

function Translate-ObjectText {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [object]$Value,
        [string]$PropertyName = "",
        [ref]$TextSegmentsTranslated
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string]) {
        $shouldTranslate = $PropertyName -eq "innerHtml" -or (Test-TranslatableProperty -PropertyName $PropertyName)
        if ($shouldTranslate -and (Test-TranslatableText -Text $Value)) {
            $textType = if ($PropertyName -eq "innerHtml") { "html" } else { "plain" }
            $context = if ($PropertyName) { "Property '$PropertyName'" } else { "" }
            $translated = Translate-Text -Text $Value -TextType $textType -Context $context
            if ($translated -ne $Value) {
                $TextSegmentsTranslated.Value++
            }
            return $translated
        }

        return $Value
    }

    if ($Value -is [array]) {
        $translatedArray = @()
        foreach ($item in $Value) {
            $translatedArray += Translate-ObjectText -Value $item -PropertyName $PropertyName -TextSegmentsTranslated $TextSegmentsTranslated
        }
        return , $translatedArray
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $translatedHash = [ordered]@{}
        foreach ($key in $Value.Keys) {
            $translatedHash[$key] = Translate-ObjectText -Value $Value[$key] -PropertyName $key -TextSegmentsTranslated $TextSegmentsTranslated
        }
        return $translatedHash
    }

    if ($Value -is [pscustomobject]) {
        $translatedObject = [ordered]@{}
        foreach ($property in $Value.PSObject.Properties) {
            $translatedObject[$property.Name] = Translate-ObjectText -Value $property.Value -PropertyName $property.Name -TextSegmentsTranslated $TextSegmentsTranslated
        }
        return $translatedObject
    }

    return $Value
}

#endregion

#region SharePoint Page Helpers

function Get-SiteIdFromUrl {
    param(
        [string]$Url
    )

    $headers = New-GraphHeaders
    $parsedUrl = [System.Uri]$Url
    $hostname = $parsedUrl.Host
    $sitePath = $parsedUrl.AbsolutePath.TrimEnd('/')

    Write-Log "Getting site ID for: $hostname`:$sitePath"
    $siteIdUrl = "https://graph.microsoft.com/v1.0/sites/$hostname`:$sitePath"
    $siteResponse = Invoke-GraphRequestWithRetry -Uri $siteIdUrl -Headers $headers -Method GET
    Write-Log "Site ID: $($siteResponse.id)"

    return $siteResponse.id
}

function Get-AllSitePages {
    param(
        [string]$SiteId
    )

    $headers = New-GraphHeaders -MetadataNone
    $pagesUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/pages?`$top=200"
    $pages = Get-GraphPagedResults -Uri $pagesUrl -Headers $headers

    return @($pages | Where-Object { $_.name -like "*.aspx" })
}

function Get-SitePageByWebUrl {
    param(
        [string]$SiteId,
        [string]$WebUrl
    )

    $targetUrl = [System.Uri]::UnescapeDataString($WebUrl).TrimEnd('/')
    $pages = Get-AllSitePages -SiteId $SiteId
    return @($pages | Where-Object {
            $_.webUrl -and
            [System.Uri]::UnescapeDataString($_.webUrl).TrimEnd('/').Equals($targetUrl, [System.StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1)
}

function Wait-SitePageByWebUrl {
    param(
        [string]$SiteId,
        [string]$WebUrl,
        [int]$MaxAttempts = 5
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $page = Get-SitePageByWebUrl -SiteId $SiteId -WebUrl $WebUrl
        if ($page.Count -gt 0) {
            return $page[0]
        }

        if ($attempt -lt $MaxAttempts) {
            Write-Log "Page not yet available through Graph Pages API. Retrying lookup ($attempt/$MaxAttempts)..." -Level WARNING
            Start-Sleep -Seconds 2
        }
    }

    throw "Could not resolve site page through Graph Pages API: $WebUrl"
}

function ConvertTo-GraphDrivePath {
    param(
        [string]$Path
    )

    $normalizedPath = $Path.Replace('\', '/').Trim('/')
    if ([string]::IsNullOrWhiteSpace($normalizedPath)) {
        return ""
    }

    $segments = @($normalizedPath.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    return (($segments | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/')
}

function Get-SiteServerRelativePath {
    $siteUri = [System.Uri]$siteUrl
    return [System.Uri]::UnescapeDataString($siteUri.AbsolutePath.TrimEnd('/'))
}

function Get-SitePagesFileServerRelativePath {
    param(
        [string]$FolderPath,
        [string]$PageName
    )

    $siteRelativePath = Get-SiteServerRelativePath
    $normalizedFolderPath = $FolderPath.Replace('\', '/').Trim('/')
    if ([string]::IsNullOrWhiteSpace($normalizedFolderPath)) {
        return "$siteRelativePath/SitePages/$PageName"
    }

    return "$siteRelativePath/SitePages/$normalizedFolderPath/$PageName"
}

function Get-ServerRelativePathFromWebUrl {
    param(
        [string]$WebUrl
    )

    return [System.Uri]::UnescapeDataString(([System.Uri]$WebUrl).AbsolutePath)
}

function ConvertTo-ODataStringLiteral {
    param(
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function ConvertTo-SharePointRestPathLiteral {
    param(
        [string]$ServerRelativePath
    )

    $normalizedPath = $ServerRelativePath.Replace('\', '/')
    $encodedPath = (($normalizedPath.Split('/') | ForEach-Object {
                if ([string]::IsNullOrEmpty($_)) {
                    ""
                }
                else {
                    [System.Uri]::EscapeDataString($_)
                }
            }) -join '/')

    return ConvertTo-ODataStringLiteral -Value $encodedPath
}

function ConvertTo-SharePointWebUrl {
    param(
        [string]$ServerRelativePath
    )

    $siteUri = [System.Uri]$siteUrl
    $normalizedPath = $ServerRelativePath.Replace('\', '/').Trim('/')
    $encodedPath = (($normalizedPath.Split('/') | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/')
    return "https://$($siteUri.Host)/$encodedPath"
}

function ConvertTo-GraphShareId {
    param(
        [string]$WebUrl
    )

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($WebUrl)
    $base64 = [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('/', '_').Replace('+', '-')
    return "u!$base64"
}

function Get-GraphDriveItemByWebUrl {
    param(
        [string]$WebUrl
    )

    $headers = New-GraphHeaders -MetadataNone
    $shareId = ConvertTo-GraphShareId -WebUrl $WebUrl
    $itemUrl = "https://graph.microsoft.com/v1.0/shares/$shareId/driveItem"
    return Invoke-GraphRequestWithRetry -Uri $itemUrl -Headers $headers -Method GET
}

function Get-SharePointFileListItemByPath {
    param(
        [string]$ServerRelativePath,
        [string[]]$Select = @("Id", "GUID", "UniqueId", "FileLeafRef", "Title")
    )

    if ([string]::IsNullOrWhiteSpace($global:sharePointToken)) {
        Get-SharePointToken
    }

    $headers = New-SharePointHeaders
    $pathLiteral = ConvertTo-SharePointRestPathLiteral -ServerRelativePath $ServerRelativePath
    $selectQuery = ($Select | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ','
    $itemUrl = "$($siteUrl.TrimEnd('/'))/_api/web/GetFileByServerRelativePath(decodedurl='$pathLiteral')/ListItemAllFields?`$select=$selectQuery"
    return Invoke-SharePointRequestWithRetry -Uri $itemUrl -Headers $headers -Method GET
}

function Get-SharePointFileListItemAllFieldsByPath {
    param(
        [string]$ServerRelativePath
    )

    if ([string]::IsNullOrWhiteSpace($global:sharePointToken)) {
        Get-SharePointToken
    }

    $headers = New-SharePointHeaders
    $pathLiteral = ConvertTo-SharePointRestPathLiteral -ServerRelativePath $ServerRelativePath
    $itemUrl = "$($siteUrl.TrimEnd('/'))/_api/web/GetFileByServerRelativePath(decodedurl='$pathLiteral')/ListItemAllFields"
    return Invoke-SharePointRequestWithRetry -Uri $itemUrl -Headers $headers -Method GET
}

function Get-SharePointFileFieldValuesForEditByPath {
    param(
        [string]$ServerRelativePath
    )

    if ([string]::IsNullOrWhiteSpace($global:sharePointToken)) {
        Get-SharePointToken
    }

    $headers = New-SharePointHeaders
    $pathLiteral = ConvertTo-SharePointRestPathLiteral -ServerRelativePath $ServerRelativePath
    $itemUrl = "$($siteUrl.TrimEnd('/'))/_api/web/GetFileByServerRelativePath(decodedurl='$pathLiteral')/ListItemAllFields/FieldValuesForEdit"
    return Invoke-SharePointRequestWithRetry -Uri $itemUrl -Headers $headers -Method GET
}

function Get-SharePointSitePageByListItemId {
    param(
        [int]$ListItemId
    )

    if ([string]::IsNullOrWhiteSpace($global:sharePointToken)) {
        Get-SharePointToken
    }

    $headers = New-SharePointHeaders
    $pageUrl = "$($siteUrl.TrimEnd('/'))/_api/sitepages/pages($ListItemId)"
    return Invoke-SharePointRequestWithRetry -Uri $pageUrl -Headers $headers -Method GET
}

function Save-SharePointSitePageByListItemId {
    param(
        [int]$ListItemId,
        [hashtable]$Fields
    )

    if ([string]::IsNullOrWhiteSpace($global:sharePointToken)) {
        Get-SharePointToken
    }

    $headers = New-SharePointHeaders
    $saveUrl = "$($siteUrl.TrimEnd('/'))/_api/sitepages/pages($ListItemId)/SavePage"
    $body = $Fields | ConvertTo-Json -Depth 100

    Write-Log "Saving page content through Site Pages SavePage endpoint for list item $ListItemId."
    Invoke-SharePointRequestWithRetry -Uri $saveUrl -Headers $headers -Method POST -Body $body | Out-Null
}

function Test-SharePointFileExists {
    param(
        [string]$ServerRelativePath
    )

    try {
        Get-SharePointFileListItemByPath -ServerRelativePath $ServerRelativePath | Out-Null
        return $true
    }
    catch {
        $statusCode = $null
        $responseDetails = [string]$_.ErrorDetails.Message
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }

        if (
            $statusCode -eq 404 -or
            ($statusCode -in @(400, 403) -and $responseDetails -match 'sharesAccessDenied|0x80070002|cannot find the file|File Not Found|not found|does not exist|invalid|System.IO.FileNotFoundException')
        ) {
            return $false
        }

        throw
    }
}

function Move-SharePointFile {
    param(
        [string]$SourceServerRelativePath,
        [string]$TargetServerRelativePath
    )

    $sourceWebUrl = ConvertTo-SharePointWebUrl -ServerRelativePath $SourceServerRelativePath
    $targetWebUrl = ConvertTo-SharePointWebUrl -ServerRelativePath $TargetServerRelativePath

    if ([string]::IsNullOrWhiteSpace($global:sharePointToken)) {
        Get-SharePointToken
    }

    $headers = New-SharePointHeaders
    $moveUrl = "$($siteUrl.TrimEnd('/'))/_api/SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=false"
    $moveBody = @{
        srcPath  = @{
            DecodedUrl = $sourceWebUrl
        }
        destPath = @{
            DecodedUrl = $targetWebUrl
        }
        options  = @{
            KeepBoth                    = $false
            ResetAuthorAndCreatedOnCopy = $false
            ShouldBypassSharedLocks     = $true
        }
    } | ConvertTo-Json -Depth 10

    Invoke-SharePointRequestWithRetry -Uri $moveUrl -Headers $headers -Method POST -Body $moveBody | Out-Null
    return [PSCustomObject]@{
        webUrl = $targetWebUrl
    }
}

function Publish-SharePointFileByPath {
    param(
        [string]$ServerRelativePath
    )

    if ([string]::IsNullOrWhiteSpace($global:sharePointToken)) {
        Get-SharePointToken
    }

    $headers = New-SharePointHeaders
    $pathLiteral = ConvertTo-SharePointRestPathLiteral -ServerRelativePath $ServerRelativePath
    $checkInUrl = "$($siteUrl.TrimEnd('/'))/_api/web/GetFileByServerRelativePath(decodedurl='$pathLiteral')/CheckIn(comment='Translated%20page',checkintype=1)"
    $publishUrl = "$($siteUrl.TrimEnd('/'))/_api/web/GetFileByServerRelativePath(decodedurl='$pathLiteral')/Publish(comment='Translated%20page')"

    $maxPublishAttempts = 12
    $checkInCompleted = $false
    $lastCheckInError = $null
    for ($attempt = 1; $attempt -le $maxPublishAttempts; $attempt++) {
        try {
            Invoke-SharePointRequestWithRetry -Uri $checkInUrl -Headers $headers -Method POST | Out-Null
            $checkInCompleted = $true
            break
        }
        catch {
            $lastCheckInError = $_
            $statusCode = $null
            $responseDetails = [string]$_.ErrorDetails.Message
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($responseDetails -match 'not checked out|not checked|no checked out|does not require check') {
                Write-Log "Page did not require check-in before publish." -Level INFO
                $checkInCompleted = $true
                break
            }

            if ($statusCode -eq 423 -and $attempt -lt $maxPublishAttempts) {
                $waitSeconds = [Math]::Min(5 * $attempt, 30)
                Write-Log "Page is locked during check-in. Retrying in $waitSeconds seconds ($attempt/$maxPublishAttempts)..." -Level WARNING
                Start-Sleep -Seconds $waitSeconds
                continue
            }

            throw
        }
    }

    if (-not $checkInCompleted) {
        throw "Could not check in page before publish because it remained locked after $maxPublishAttempts attempts. Last error: $($lastCheckInError.Exception.Message)"
    }

    for ($attempt = 1; $attempt -le $maxPublishAttempts; $attempt++) {
        try {
            Invoke-SharePointRequestWithRetry -Uri $publishUrl -Headers $headers -Method POST | Out-Null
            Write-Log "Published SharePoint page: $ServerRelativePath" -Level SUCCESS
            return
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($statusCode -eq 423 -and $attempt -lt $maxPublishAttempts) {
                $waitSeconds = [Math]::Min(5 * $attempt, 30)
                Write-Log "Page is locked during publish. Retrying in $waitSeconds seconds ($attempt/$maxPublishAttempts)..." -Level WARNING
                Start-Sleep -Seconds $waitSeconds
                continue
            }

            throw
        }
    }

    throw "Could not publish page after $maxPublishAttempts attempts."
}

function Set-SharePointPageFieldsByValidateUpdate {
    param(
        [string]$ServerRelativePath,
        [hashtable]$Fields
    )

    if ([string]::IsNullOrWhiteSpace($global:sharePointToken)) {
        Get-SharePointToken
    }

    $headers = New-SharePointHeaders
    $pathLiteral = ConvertTo-SharePointRestPathLiteral -ServerRelativePath $ServerRelativePath
    $updateUrl = "$($siteUrl.TrimEnd('/'))/_api/web/GetFileByServerRelativePath(decodedurl='$pathLiteral')/ListItemAllFields/ValidateUpdateListItem()"
    $formValues = @()
    foreach ($key in $Fields.Keys) {
        if ($null -ne $Fields[$key]) {
            $formValues += @{
                FieldName  = $key
                FieldValue = [string]$Fields[$key]
            }
        }
    }

    $body = @{
        formValues         = $formValues
        bNewDocumentUpdate = $true
    } | ConvertTo-Json -Depth 20

    Write-Log "Updating Site Pages fields with ValidateUpdateListItem: $($Fields.Keys -join ', ')"
    Invoke-SharePointRequestWithRetry -Uri $updateUrl -Headers $headers -Method POST -Body $body | Out-Null
}

function Get-NestedPropertyValue {
    param(
        [object]$Value,
        [string[]]$PropertyNames
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($propertyName in $PropertyNames) {
            if ($Value.Contains($propertyName) -and -not [string]::IsNullOrWhiteSpace([string]($Value[$propertyName]))) {
                return $Value[$propertyName]
            }
        }

        foreach ($key in $Value.Keys) {
            $nestedValue = Get-NestedPropertyValue -Value $Value[$key] -PropertyNames $PropertyNames
            if ($null -ne $nestedValue) {
                return $nestedValue
            }
        }
    }
    elseif ($Value -is [array]) {
        foreach ($item in $Value) {
            $nestedValue = Get-NestedPropertyValue -Value $item -PropertyNames $PropertyNames
            if ($null -ne $nestedValue) {
                return $nestedValue
            }
        }
    }
    elseif ($Value -is [pscustomobject]) {
        foreach ($propertyName in $PropertyNames) {
            $property = $Value.PSObject.Properties[$propertyName]
            if ($property -and -not [string]::IsNullOrWhiteSpace([string]($property.Value))) {
                return $property.Value
            }
        }

        foreach ($property in $Value.PSObject.Properties) {
            $nestedValue = Get-NestedPropertyValue -Value $property.Value -PropertyNames $PropertyNames
            if ($null -ne $nestedValue) {
                return $nestedValue
            }
        }
    }

    return $null
}

function Get-ImageWebPartDetails {
    param(
        [object]$WebPart
    )

    $imageUrl = Get-NestedPropertyValue -Value $WebPart -PropertyNames @(
        "imageSource",
        "imageUrl",
        "serverRelativeUrl",
        "serverRelativeImageUrl",
        "originalImageUrl",
        "url"
    )

    if ($imageUrl -and $imageUrl -match '^https?://') {
        $siteUri = [System.Uri]$siteUrl
        if (([System.Uri]$imageUrl).Host -ne $siteUri.Host) {
            $imageUrl = $null
        }
    }

    if ($imageUrl -and $imageUrl -notmatch '^https?://' -and $imageUrl -notmatch '^/') {
        $imageUrl = $null
    }

    return [PSCustomObject]@{
        ImageUrl        = $imageUrl
        Caption         = Get-NestedPropertyValue -Value $WebPart -PropertyNames @("captionText", "caption")
        AlternativeText = Get-NestedPropertyValue -Value $WebPart -PropertyNames @("altText", "alternativeText")
        Width           = Get-NestedPropertyValue -Value $WebPart -PropertyNames @("imgWidth", "imageWidth", "width")
        Height          = Get-NestedPropertyValue -Value $WebPart -PropertyNames @("imgHeight", "imageHeight", "height")
    }
}

function New-RestCanvasControlPosition {
    param(
        [int]$SectionIndex,
        [int]$ColumnIndex,
        [int]$ControlIndex
    )

    return [ordered]@{
        zoneIndex    = $SectionIndex
        sectionIndex = $ColumnIndex
        controlIndex = $ControlIndex
        layoutIndex  = 1
    }
}

function New-RestCanvasTextControl {
    param(
        [string]$Html,
        [int]$SectionIndex,
        [int]$ColumnIndex,
        [int]$ControlIndex
    )

    return [ordered]@{
        controlType = 4
        id          = [guid]::NewGuid().ToString()
        position    = New-RestCanvasControlPosition -SectionIndex $SectionIndex -ColumnIndex $ColumnIndex -ControlIndex $ControlIndex
        innerHTML   = $Html
    }
}

function New-RestCanvasWebPartControl {
    param(
        [object]$WebPart,
        [int]$SectionIndex,
        [int]$ColumnIndex,
        [int]$ControlIndex
    )

    $webPartType = [string]$WebPart["webPartType"]
    if ([string]::IsNullOrWhiteSpace($webPartType) -or $webPartType -eq "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788") {
        return $null
    }

    $webPartData = [ordered]@{
        id                     = $webPartType
        instanceId             = [guid]::NewGuid().ToString()
        title                  = ""
        description            = ""
        dataVersion            = "1.0"
        properties             = [ordered]@{}
        serverProcessedContent = [ordered]@{
            htmlStrings          = [ordered]@{}
            searchablePlainTexts = [ordered]@{}
            imageSources         = [ordered]@{}
            links                = [ordered]@{}
        }
    }

    if ($WebPart.Contains("data")) {
        $data = $WebPart["data"]
        if ($data -is [System.Collections.IDictionary]) {
            foreach ($key in @("title", "description", "dataVersion", "properties", "serverProcessedContent")) {
                if ($data.Contains($key) -and $null -ne $data[$key]) {
                    $webPartData[$key] = $data[$key]
                }
            }
        }
    }

    return [ordered]@{
        controlType = 3
        id          = [guid]::NewGuid().ToString()
        position    = New-RestCanvasControlPosition -SectionIndex $SectionIndex -ColumnIndex $ColumnIndex -ControlIndex $ControlIndex
        webPartId   = $webPartType
        webPartData = $webPartData
    }
}

function Convert-GraphPayloadToRestCanvasContent {
    param(
        [object]$Payload
    )

    $controls = @()
    $sectionIndex = 1
    if ($Payload["canvasLayout"] -and $Payload["canvasLayout"].Contains("horizontalSections")) {
        foreach ($section in $Payload["canvasLayout"]["horizontalSections"]) {
            $columnIndex = 1
            foreach ($column in @($section["columns"])) {
                $controlIndex = 1
                foreach ($webPart in @($column["webparts"])) {
                    if ($webPart.Contains("innerHtml") -and -not [string]::IsNullOrWhiteSpace([string]$webPart["innerHtml"])) {
                        $controls += New-RestCanvasTextControl -Html ([string]$webPart["innerHtml"]) `
                            -SectionIndex $sectionIndex -ColumnIndex $columnIndex -ControlIndex $controlIndex
                    }
                    elseif ($webPart.Contains("webPartType")) {
                        $control = New-RestCanvasWebPartControl -WebPart $webPart `
                            -SectionIndex $sectionIndex -ColumnIndex $columnIndex -ControlIndex $controlIndex
                        if ($null -ne $control) {
                            $controls += $control
                        }
                    }

                    $controlIndex++
                }

                $columnIndex++
            }

            $sectionIndex++
        }
    }

    return ($controls | ConvertTo-Json -Depth 100 -Compress)
}

function Translate-SharePointJsonField {
    param(
        [string]$Json,
        [ref]$TextSegmentsTranslated
    )

    if ([string]::IsNullOrWhiteSpace($Json)) {
        return ""
    }

    try {
        $jsonObject = $Json | ConvertFrom-Json -Depth 100
    }
    catch {
        if ($Json.TrimStart().StartsWith("<")) {
            $translatedHtml = Translate-Text -Text $Json -TextType html -Context "SharePoint CanvasContent1 HTML"
            if ($translatedHtml -ne $Json) {
                $TextSegmentsTranslated.Value++
            }
            return $translatedHtml
        }

        Write-Log "Could not parse SharePoint page JSON field for translation. Keeping original value. Error: $($_.Exception.Message)" -Level WARNING
        return $Json
    }

    $translatedObject = Translate-ObjectText -Value $jsonObject -TextSegmentsTranslated $TextSegmentsTranslated
    return ($translatedObject | ConvertTo-Json -Depth 100 -Compress)
}

function New-LayoutWebpartsContent {
    param(
        [object]$Payload
    )

    if ($Payload["titleArea"]) {
        $titleArea = $Payload["titleArea"]
        $layoutData = [ordered]@{
            id                     = "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"
            instanceId             = [guid]::NewGuid().ToString()
            title                  = "Title area"
            description            = "Title area"
            serverProcessedContent = [ordered]@{
                htmlStrings          = [ordered]@{}
                searchablePlainTexts = [ordered]@{ title = [string]$Payload["title"] }
                imageSources         = [ordered]@{}
                links                = [ordered]@{}
            }
            dataVersion            = "1.4"
            properties             = $titleArea
        }
        return ($layoutData | ConvertTo-Json -Depth 100 -Compress)
    }

    return ""
}

function New-RestSubfolderTranslatedPage {
    param(
        [object]$Page,
        [object]$SourceContent,
        [string]$TargetPageName,
        [ref]$TextSegmentsTranslated,
        [switch]$Publish
    )

    $sourceFolderPath = Get-FolderPathFromWebUrl -WebUrl $Page.webUrl
    $targetServerRelativePath = Get-SitePagesFileServerRelativePath -FolderPath $sourceFolderPath -PageName $TargetPageName
    $targetFolderPath = [System.IO.Path]::GetDirectoryName($targetServerRelativePath).Replace('\', '/')
    $targetFileName = [System.IO.Path]::GetFileName($targetServerRelativePath)
    $payload = New-TranslatedPagePayload -SourcePageContent $SourceContent -NewPageName $TargetPageName `
        -TextSegmentsTranslated $TextSegmentsTranslated
    $targetTitle = $payload["title"]
    if ([string]::IsNullOrWhiteSpace($targetTitle)) {
        $targetTitle = [System.IO.Path]::GetFileNameWithoutExtension($TargetPageName)
    }

    if ([string]::IsNullOrWhiteSpace($global:sharePointToken)) {
        Get-SharePointToken
    }

    $headers = New-SharePointHeaders
    $folderLiteral = ConvertTo-SharePointRestPathLiteral -ServerRelativePath $targetFolderPath
    $targetPathLiteral = ConvertTo-SharePointRestPathLiteral -ServerRelativePath $targetServerRelativePath
    $addUrl = "$($siteUrl.TrimEnd('/'))/_api/web/GetFolderByServerRelativePath(decodedurl='$folderLiteral')/Files/AddTemplateFile(urlOfFile='$targetPathLiteral',templateFileType=3)"

    Write-Log "Creating translated subfolder page with SharePoint REST: $targetServerRelativePath"
    Invoke-SharePointRequestWithRetry -Uri $addUrl -Headers $headers -Method POST | Out-Null
    Start-Sleep -Seconds 3

    $sourceServerRelativePath = Get-ServerRelativePathFromWebUrl -WebUrl $Page.webUrl
    $sourceItem = Get-SharePointFileListItemAllFieldsByPath -ServerRelativePath $sourceServerRelativePath
    $sourceEditFields = Get-SharePointFileFieldValuesForEditByPath -ServerRelativePath $sourceServerRelativePath
    $sourceFieldNames = @($sourceItem.PSObject.Properties.Name | Where-Object { $_ -match 'Canvas|Layout|Content|Banner|Description|Title' })
    Write-Log "Source REST page fields available: $($sourceFieldNames -join ', ')"
    Write-Log "Source ListItemAllFields CanvasContent1 length: $(([string]$sourceItem.CanvasContent1).Length)"
    Write-Log "Source ListItemAllFields LayoutWebpartsContent length: $(([string]$sourceItem.LayoutWebpartsContent).Length)"
    Write-Log "Source FieldValuesForEdit CanvasContent1 length: $(([string]$sourceEditFields.CanvasContent1).Length)"
    Write-Log "Source FieldValuesForEdit LayoutWebpartsContent length: $(([string]$sourceEditFields.LayoutWebpartsContent).Length)"
    $sourceCanvasJson = if (-not [string]::IsNullOrWhiteSpace([string]$sourceEditFields.CanvasContent1)) { [string]$sourceEditFields.CanvasContent1 } else { [string]$sourceItem.CanvasContent1 }
    $sourceLayoutJson = if (-not [string]::IsNullOrWhiteSpace([string]$sourceEditFields.LayoutWebpartsContent)) { [string]$sourceEditFields.LayoutWebpartsContent } else { [string]$sourceItem.LayoutWebpartsContent }
    $canvasContent = Translate-SharePointJsonField -Json $sourceCanvasJson -TextSegmentsTranslated $TextSegmentsTranslated
    $layoutWebpartsContent = Translate-SharePointJsonField -Json $sourceLayoutJson -TextSegmentsTranslated $TextSegmentsTranslated
    if ([string]::IsNullOrWhiteSpace($canvasContent)) {
        Write-Log "Source CanvasContent1 was not available through REST. Falling back to Graph payload canvas conversion." -Level WARNING
        $canvasContent = Convert-GraphPayloadToRestCanvasContent -Payload $payload
        Write-Log "Generated fallback CanvasContent1 length: $(([string]$canvasContent).Length)"
    }
    if ([string]::IsNullOrWhiteSpace($layoutWebpartsContent)) {
        $layoutWebpartsContent = New-LayoutWebpartsContent -Payload $payload
        Write-Log "Generated fallback LayoutWebpartsContent length: $(([string]$layoutWebpartsContent).Length)"
    }

    $item = Get-SharePointFileListItemByPath -ServerRelativePath $targetServerRelativePath
    $pathLiteral = ConvertTo-SharePointRestPathLiteral -ServerRelativePath $targetServerRelativePath
    $itemUrl = "$($siteUrl.TrimEnd('/'))/_api/web/GetFileByServerRelativePath(decodedurl='$pathLiteral')/ListItemAllFields"
    $mergeHeaders = New-SharePointHeaders
    $mergeHeaders["IF-MATCH"] = "*"
    $mergeHeaders["X-HTTP-Method"] = "MERGE"
    $bodyObject = @{
        Title                   = $targetTitle
        CanvasContent1          = $canvasContent
        LayoutWebpartsContent   = $layoutWebpartsContent
        ClientSideApplicationId = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec"
        PageLayoutType          = "Article"
        PromotedState           = 0
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$sourceItem.BannerImageUrl)) {
        $bodyObject["BannerImageUrl"] = [string]$sourceItem.BannerImageUrl
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$sourceItem.Description)) {
        $bodyObject["Description"] = Translate-Text -Text ([string]$sourceItem.Description) -TextType plain
    }

    $body = $bodyObject | ConvertTo-Json -Depth 100

    Invoke-SharePointRequestWithRetry -Uri $itemUrl -Headers $mergeHeaders -Method POST -Body $body | Out-Null
    try {
        Save-SharePointSitePageByListItemId -ListItemId ([int]$item.Id) -Fields $bodyObject
    }
    catch {
        Write-Log "Site Pages SavePage endpoint did not accept the payload: $($_.Exception.Message)" -Level WARNING
    }

    $targetItemAfterUpdate = Get-SharePointFileListItemAllFieldsByPath -ServerRelativePath $targetServerRelativePath
    Write-Log "Target ListItemAllFields CanvasContent1 length after update: $(([string]$targetItemAfterUpdate.CanvasContent1).Length)"
    Write-Log "Target ListItemAllFields LayoutWebpartsContent length after update: $(([string]$targetItemAfterUpdate.LayoutWebpartsContent).Length)"

    if ([string]::IsNullOrWhiteSpace([string]$targetItemAfterUpdate.CanvasContent1)) {
        Set-SharePointPageFieldsByValidateUpdate -ServerRelativePath $targetServerRelativePath -Fields @{
            Title                   = $targetTitle
            CanvasContent1          = $canvasContent
            LayoutWebpartsContent   = $layoutWebpartsContent
            ClientSideApplicationId = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec"
            PageLayoutType          = "Article"
            PromotedState           = "0"
        }

        $targetItemAfterValidateUpdate = Get-SharePointFileListItemAllFieldsByPath -ServerRelativePath $targetServerRelativePath
        Write-Log "Target ListItemAllFields CanvasContent1 length after ValidateUpdateListItem: $(([string]$targetItemAfterValidateUpdate.CanvasContent1).Length)"
        Write-Log "Target ListItemAllFields LayoutWebpartsContent length after ValidateUpdateListItem: $(([string]$targetItemAfterValidateUpdate.LayoutWebpartsContent).Length)"
    }

    if ($Publish) {
        Start-Sleep -Seconds 5
        Publish-SharePointFileByPath -ServerRelativePath $targetServerRelativePath
    }

    $siteUri = [System.Uri]$siteUrl
    return [PSCustomObject]@{
        id     = $(if ($item.GUID) { ([string]$item.GUID).Trim('{}') } else { $Page.id })
        name   = $TargetPageName
        webUrl = "https://$($siteUri.Host)$targetServerRelativePath"
    }
}

function Normalize-PageReferencePath {
    param(
        [string]$Reference
    )

    if ([string]::IsNullOrWhiteSpace($Reference)) {
        return ""
    }

    $referencePath = $Reference.Trim()
    if ($referencePath -match '^https?://') {
        $uri = [System.Uri]$referencePath
        $referencePath = $uri.AbsolutePath
    }

    $referencePath = [System.Uri]::UnescapeDataString($referencePath).Replace('\', '/').Trim('/')
    if ($referencePath -match '(?i)(?:^|/)SitePages/(.+)$') {
        $referencePath = $matches[1]
    }

    return $referencePath.Trim('/')
}

function Get-PageFileNameFromReference {
    param(
        [string]$Reference
    )

    $normalizedPath = Normalize-PageReferencePath -Reference $Reference
    if ([string]::IsNullOrWhiteSpace($normalizedPath)) {
        return ""
    }

    return @($normalizedPath.Split('/'))[-1]
}

function Get-FolderPathFromPageReference {
    param(
        [string]$Reference
    )

    $normalizedPath = Normalize-PageReferencePath -Reference $Reference
    $segments = @($normalizedPath.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($segments.Count -le 1) {
        return ""
    }

    return "/$($segments[0..($segments.Count - 2)] -join '/')"
}

function Get-SitePagesDriveId {
    param(
        [string]$SiteId
    )

    if ($script:sitePagesDriveIds.ContainsKey($SiteId)) {
        return $script:sitePagesDriveIds[$SiteId]
    }

    $headers = New-GraphHeaders -MetadataNone
    $drivesUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives?`$select=id,name,webUrl,driveType&`$top=200"
    $drives = Get-GraphPagedResults -Uri $drivesUrl -Headers $headers
    $sitePagesDrive = @($drives | Where-Object {
            $_.name -eq "Site Pages" -or
            $_.name -eq "SitePages" -or
            $_.webUrl -match '/SitePages/?$'
        } | Select-Object -First 1)

    if ($sitePagesDrive.Count -gt 0) {
        $script:sitePagesDriveIds[$SiteId] = $sitePagesDrive[0].id
        Write-Log "Site Pages drive ID: $($sitePagesDrive[0].id)"
        return $sitePagesDrive[0].id
    }

    $listsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists?`$select=id,displayName,webUrl&`$top=200"
    $lists = Get-GraphPagedResults -Uri $listsUrl -Headers $headers
    $sitePagesList = @($lists | Where-Object {
            $_.displayName -eq "Site Pages" -or
            $_.displayName -eq "SitePages" -or
            $_.webUrl -match '/SitePages/?$'
        } | Select-Object -First 1)

    if ($sitePagesList.Count -eq 0) {
        $availableDrives = ($drives | ForEach-Object { "$($_.name) [$($_.webUrl)]" }) -join "; "
        throw "Could not find the Site Pages library for site '$SiteId'. Available drives: $availableDrives"
    }

    $driveUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$($sitePagesList[0].id)/drive"
    $drive = Invoke-GraphRequestWithRetry -Uri $driveUrl -Headers $headers -Method GET
    $script:sitePagesDriveIds[$SiteId] = $drive.id
    return $drive.id
}

function Get-SitePagesDriveItemByPath {
    param(
        [string]$SiteId,
        [string]$RelativePath,
        [switch]$IncludeListItem
    )

    $driveId = Get-SitePagesDriveId -SiteId $SiteId
    $encodedPath = ConvertTo-GraphDrivePath -Path $RelativePath
    if ([string]::IsNullOrWhiteSpace($encodedPath)) {
        throw "A Site Pages drive relative path is required."
    }

    $expandQuery = ""
    if ($IncludeListItem) {
        $expandQuery = "?`$expand=listItem(`$expand=fields)"
    }

    $headers = New-GraphHeaders -MetadataNone
    $itemUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/$($encodedPath):$expandQuery"
    return Invoke-GraphRequestWithRetry -Uri $itemUrl -Headers $headers -Method GET
}

function Test-SitePagesDrivePathExists {
    param(
        [string]$SiteId,
        [string]$RelativePath
    )

    try {
        Get-SitePagesDriveItemByPath -SiteId $SiteId -RelativePath $RelativePath | Out-Null
        return $true
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }

        if ($statusCode -eq 404) {
            return $false
        }

        throw
    }
}

function Find-SitePagesDriveItemsByName {
    param(
        [string]$SiteId,
        [string]$FileName
    )

    $driveId = Get-SitePagesDriveId -SiteId $SiteId
    $headers = New-GraphHeaders -MetadataNone
    $escapedFileName = $FileName.Replace("'", "''")
    $searchUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/root/search(q='$escapedFileName')?`$top=200"
    $items = Get-GraphPagedResults -Uri $searchUrl -Headers $headers

    return @($items | Where-Object {
            $_.file -and
            $_.name -eq $FileName -and
            $_.name -like "*.aspx"
        })
}

function Get-SitePagesDrivePageItems {
    param(
        [string]$SiteId,
        [string]$FolderItemId = ""
    )

    $driveId = Get-SitePagesDriveId -SiteId $SiteId
    $headers = New-GraphHeaders -MetadataNone
    if ([string]::IsNullOrWhiteSpace($FolderItemId)) {
        $childrenUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/root/children?`$top=200"
    }
    else {
        $childrenUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$FolderItemId/children?`$top=200"
    }

    $children = Get-GraphPagedResults -Uri $childrenUrl -Headers $headers
    $pageItems = @()
    foreach ($child in $children) {
        if ($child.folder) {
            $pageItems += Get-SitePagesDrivePageItems -SiteId $SiteId -FolderItemId $child.id
        }
        elseif ($child.file -and $child.name -like "*.aspx") {
            $pageItems += $child
        }
    }

    return $pageItems
}

function Convert-DriveItemToSitePageReference {
    param(
        [string]$SiteId,
        [object]$DriveItem
    )

    if ($null -eq $DriveItem.listItem -or $null -eq $DriveItem.listItem.fields) {
        $driveId = Get-SitePagesDriveId -SiteId $SiteId
        $headers = New-GraphHeaders -MetadataNone
        $itemUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$($DriveItem.id)?`$expand=listItem(`$expand=fields)"
        $DriveItem = Invoke-GraphRequestWithRetry -Uri $itemUrl -Headers $headers -Method GET
    }

    $pageId = $DriveItem.sharepointIds.listItemUniqueId
    if ([string]::IsNullOrWhiteSpace($pageId) -and $DriveItem.listItem.fields.UniqueId) {
        $pageId = $DriveItem.listItem.fields.UniqueId
    }

    if ([string]::IsNullOrWhiteSpace($pageId)) {
        throw "Could not resolve a site page ID for '$($DriveItem.webUrl)'."
    }

    return [PSCustomObject]@{
        id     = $pageId
        name   = $DriveItem.name
        title  = $(if ($DriveItem.listItem.fields.Title) { $DriveItem.listItem.fields.Title } else { [System.IO.Path]::GetFileNameWithoutExtension($DriveItem.name) })
        webUrl = $DriveItem.webUrl
    }
}

function Resolve-PageReference {
    param(
        [string]$SiteId,
        [array]$Pages,
        [string]$Reference
    )

    if ([string]::IsNullOrWhiteSpace($Reference)) {
        return $Pages
    }

    $candidate = Get-PageFileNameFromReference -Reference $Reference
    $candidateWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($candidate)
    $candidateFolderPath = Get-FolderPathFromPageReference -Reference $Reference

    $matches = @($Pages | Where-Object {
            $pageFolderPath = Get-FolderPathFromWebUrl -WebUrl $_.webUrl
            $folderMatches = [string]::IsNullOrWhiteSpace($candidateFolderPath) -or
            $pageFolderPath.Equals($candidateFolderPath, [System.StringComparison]::OrdinalIgnoreCase)

            $folderMatches -and (
                $_.name -eq $candidate -or
                $_.name -eq "$candidate.aspx" -or
                $_.title -eq $candidate -or
                $_.title -eq $candidateWithoutExtension
            )
        })

    if ($matches.Count -eq 0) {
        $matches = @($Pages | Where-Object {
                $pageFolderPath = Get-FolderPathFromWebUrl -WebUrl $_.webUrl
                $folderMatches = [string]::IsNullOrWhiteSpace($candidateFolderPath) -or
                $pageFolderPath.Equals($candidateFolderPath, [System.StringComparison]::OrdinalIgnoreCase)

                $folderMatches -and (
                    $_.name -like $candidate -or
                    $_.title -like $candidate -or
                    $_.name -like "$candidateWithoutExtension.aspx" -or
                    $_.title -like $candidateWithoutExtension
                )
            })
    }

    if ($matches.Count -eq 0) {
        $candidateRelativePath = Normalize-PageReferencePath -Reference $Reference
        if ($candidateRelativePath -match '/') {
            $driveItem = Get-SitePagesDriveItemByPath -SiteId $SiteId -RelativePath $candidateRelativePath -IncludeListItem
            return Convert-DriveItemToSitePageReference -SiteId $SiteId -DriveItem $driveItem
        }

        $driveMatches = @(Find-SitePagesDriveItemsByName -SiteId $SiteId -FileName $candidate)
        if ($driveMatches.Count -eq 1) {
            return Convert-DriveItemToSitePageReference -SiteId $SiteId -DriveItem $driveMatches[0]
        }

        if ($driveMatches.Count -gt 1) {
            $names = ($driveMatches | ForEach-Object { "$($_.name) [$($_.webUrl)]" }) -join "; "
            throw "Page reference '$Reference' matched multiple pages in Site Pages folders: $names"
        }
    }

    if ($matches.Count -eq 0) {
        throw "Could not find a SharePoint page matching '$Reference'."
    }

    if ($matches.Count -gt 1) {
        $names = ($matches | ForEach-Object { "$($_.name) [$($_.title)]" }) -join "; "
        throw "Page reference '$Reference' matched multiple pages: $names"
    }

    return $matches
}

function Get-SitePageContent {
    param(
        [string]$SiteId,
        [string]$PageId
    )

    $headers = New-GraphHeaders -MetadataNone
    $pageUrl = "https://graph.microsoft.com/beta/sites/$SiteId/pages/$PageId/microsoft.graph.sitePage?`$expand=canvasLayout"
    return Invoke-GraphRequestWithRetry -Uri $pageUrl -Headers $headers -Method GET
}

function Remove-ReadOnlyPageProperties {
    param(
        [object]$Value
    )

    $readOnlyNames = @(
        '@odata.context',
        '@odata.etag',
        'eTag',
        'webUrl',
        'createdBy',
        'createdDateTime',
        'lastModifiedBy',
        'lastModifiedDateTime',
        'parentReference',
        'contentType',
        'publishingState',
        'reactions',
        'thumbnailWebUrl',
        'customContentDropSupport'
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [array]) {
        $cleanArray = @()
        foreach ($item in $Value) {
            $cleanArray += Remove-ReadOnlyPageProperties -Value $item
        }
        return , $cleanArray
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $cleanHash = [ordered]@{}
        foreach ($key in $Value.Keys) {
            if ($readOnlyNames -notcontains $key) {
                $cleanHash[$key] = Remove-ReadOnlyPageProperties -Value $Value[$key]
            }
        }
        return $cleanHash
    }

    if ($Value -is [pscustomobject]) {
        $cleanObject = [ordered]@{}
        foreach ($property in $Value.PSObject.Properties) {
            if ($readOnlyNames -notcontains $property.Name) {
                $cleanObject[$property.Name] = Remove-ReadOnlyPageProperties -Value $property.Value
            }
        }
        return $cleanObject
    }

    return $Value
}

function Reset-WebPartIds {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return
    }

    if ($Value -is [array]) {
        foreach ($item in $Value) {
            Reset-WebPartIds -Value $item
        }
        return
    }

    if ($Value -is [System.Collections.IDictionary]) {
        if ($Value.Contains('webparts') -and $Value['webparts']) {
            foreach ($webPart in $Value['webparts']) {
                if ($webPart -is [System.Collections.IDictionary] -and $webPart.Contains('id')) {
                    $webPart['id'] = [guid]::NewGuid().ToString()
                }
            }
        }

        foreach ($key in $Value.Keys) {
            Reset-WebPartIds -Value $Value[$key]
        }
        return
    }
}

function Remove-EmptyNestedTitleProperties {
    param(
        [object]$Value,
        [bool]$IsTopLevel = $false
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [array]) {
        $cleanArray = @()
        foreach ($item in $Value) {
            $cleanArray += Remove-EmptyNestedTitleProperties -Value $item
        }
        return , $cleanArray
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $cleanHash = [ordered]@{}
        foreach ($key in $Value.Keys) {
            if (-not $IsTopLevel -and $key -eq "title" -and [string]::IsNullOrWhiteSpace([string]($Value[$key]))) {
                continue
            }

            $cleanHash[$key] = Remove-EmptyNestedTitleProperties -Value $Value[$key]
        }
        return $cleanHash
    }

    if ($Value -is [pscustomobject]) {
        $cleanObject = [ordered]@{}
        foreach ($property in $Value.PSObject.Properties) {
            if (-not $IsTopLevel -and $property.Name -eq "title" -and [string]::IsNullOrWhiteSpace([string]($property.Value))) {
                continue
            }

            $cleanObject[$property.Name] = Remove-EmptyNestedTitleProperties -Value $property.Value
        }
        return $cleanObject
    }

    return $Value
}

function Get-TranslatedPageName {
    param(
        [string]$SourceName,
        [array]$ExistingPages,
        [string]$SourceFolderPath = ""
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourceName)
    $extension = [System.IO.Path]::GetExtension($SourceName)
    if ([string]::IsNullOrWhiteSpace($extension)) {
        $extension = ".aspx"
    }

    $candidate = "$baseName$OutputNameSuffix$extension"
    $existingNames = @($ExistingPages | Where-Object {
            $existingFolderPath = Get-FolderPathFromWebUrl -WebUrl $_.webUrl
            if ([string]::IsNullOrWhiteSpace($SourceFolderPath)) {
                [string]::IsNullOrWhiteSpace($existingFolderPath)
            }
            else {
                $existingFolderPath.Equals($SourceFolderPath, [System.StringComparison]::OrdinalIgnoreCase)
            }
        } | ForEach-Object { $_.name })
    
    Write-Log "Checking for name conflicts. Looking for: '$candidate'" -Level INFO
    Write-Log "Source page folder path: $(if ([string]::IsNullOrWhiteSpace($SourceFolderPath)) { '(root)' } else { $SourceFolderPath })" -Level INFO

    if ($existingNames -notcontains $candidate) {
        Write-Log "Name '$candidate' is available" -Level INFO
        return $candidate
    }

    # Name exists, generate a unique one with timestamp to avoid conflicts
    Write-Log "Page name '$candidate' already exists. Generating unique name..." -Level WARNING
    $uniqueSuffix = Get-Date -Format "yyyyMMddHHmmss"
    $uniqueName = "$baseName$OutputNameSuffix-$uniqueSuffix$extension"
    Write-Log "Using unique page name: $uniqueName" -Level INFO
    return $uniqueName
}

function Get-FolderPathFromWebUrl {
    param(
        [string]$WebUrl
    )

    if ([string]::IsNullOrWhiteSpace($WebUrl)) {
        return ""
    }

    $path = $WebUrl
    if ($WebUrl -match '^https?://') {
        $path = ([System.Uri]$WebUrl).AbsolutePath
    }

    $path = [System.Uri]::UnescapeDataString($path).Replace('\', '/')
    if ($path -match '(?i)/SitePages/(.+)/[^/]+\.aspx$') {
        return "/$($matches[1])"
    }
    
    return ""
}

function Move-PageToFolder {
    param(
        [string]$SiteId,
        [string]$PageName,
        [string]$FolderPath,
        [string]$TargetPageName = ""
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath) -or $FolderPath -eq "/") {
        # Already at root, no move needed
        return
    }

    if ([string]::IsNullOrWhiteSpace($TargetPageName)) {
        $TargetPageName = $PageName
    }

    try {
        Write-Log "Moving page '$PageName' to folder '$FolderPath' as '$TargetPageName'..."

        $sourceServerRelativePath = Get-SitePagesFileServerRelativePath -FolderPath "" -PageName $PageName
        $targetServerRelativePath = Get-SitePagesFileServerRelativePath -FolderPath $FolderPath -PageName $TargetPageName
        $sourceLookupAttempts = 5
        for ($attempt = 1; $attempt -le $sourceLookupAttempts; $attempt++) {
            if (Test-SharePointFileExists -ServerRelativePath $sourceServerRelativePath) {
                break
            }

            if ($attempt -lt $sourceLookupAttempts) {
                Write-Log "Created page not yet discoverable at '$sourceServerRelativePath'. Retrying move lookup ($attempt/$sourceLookupAttempts)..." -Level WARNING
                Start-Sleep -Seconds 2
                continue
            }

            throw "Created page is not discoverable at '$sourceServerRelativePath'."
        }

        if (Test-SharePointFileExists -ServerRelativePath $targetServerRelativePath) {
            throw "Target page already exists at '$targetServerRelativePath'."
        }

        $movedItem = Move-SharePointFile -SourceServerRelativePath $sourceServerRelativePath -TargetServerRelativePath $targetServerRelativePath
        Write-Log "Successfully moved page to folder: $FolderPath" -Level SUCCESS
        return $movedItem
    }
    catch {
        Write-Log "Failed to move page to folder: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function New-TranslatedPagePayload {
    param(
        [object]$SourcePageContent,
        [string]$NewPageName,
        [ref]$TextSegmentsTranslated
    )

    # Handle empty or null titles
    $sourceTitle = $SourcePageContent.title
    if ([string]::IsNullOrWhiteSpace($sourceTitle)) {
        $sourceTitle = [System.IO.Path]::GetFileNameWithoutExtension($SourcePageContent.name)
        Write-Log "Source page title was empty. Using page name without extension: $sourceTitle" -Level WARNING
    }

    $translatedTitle = Translate-Text -Text $sourceTitle -TextType plain
    if ([string]::IsNullOrWhiteSpace($translatedTitle)) {
        Write-Log "Translated page title was empty. Falling back to source title: $sourceTitle" -Level WARNING
        $translatedTitle = $sourceTitle
    }

    if ($translatedTitle -ne $sourceTitle) {
        $TextSegmentsTranslated.Value++
    }

    $payload = [ordered]@{
        "@odata.type"        = "#microsoft.graph.sitePage"
        name                 = $NewPageName
        title                = $translatedTitle
        pageLayout           = $(if ($SourcePageContent.pageLayout) { $SourcePageContent.pageLayout } else { "article" })
        showComments         = $SourcePageContent.showComments
        showRecommendedPages = $SourcePageContent.showRecommendedPages
    }

    if ($SourcePageContent.description) {
        $translatedDescription = Translate-Text -Text $SourcePageContent.description -TextType plain
        if ($translatedDescription -ne $SourcePageContent.description) {
            $TextSegmentsTranslated.Value++
        }
        $payload["description"] = $translatedDescription
    }

    if ($SourcePageContent.titleArea) {
        $payload["titleArea"] = Remove-ReadOnlyPageProperties -Value (Translate-ObjectText -Value $SourcePageContent.titleArea -TextSegmentsTranslated $TextSegmentsTranslated)
        if ($payload["titleArea"].Contains("title")) {
            $payload["titleArea"]["title"] = $translatedTitle
        }
    }

    if ($SourcePageContent.canvasLayout) {
        $translatedCanvas = Translate-ObjectText -Value $SourcePageContent.canvasLayout -TextSegmentsTranslated $TextSegmentsTranslated
        $payload["canvasLayout"] = Remove-ReadOnlyPageProperties -Value $translatedCanvas
        Reset-WebPartIds -Value $payload["canvasLayout"]
    }

    return $payload
}

function Publish-SitePage {
    param(
        [string]$SiteId,
        [string]$PageId
    )

    $headers = New-GraphHeaders
    $publishUrl = "https://graph.microsoft.com/beta/sites/$SiteId/pages/$PageId/microsoft.graph.sitePage/publish"
    Invoke-GraphRequestWithRetry -Uri $publishUrl -Headers $headers -Method POST -Body "{}" | Out-Null
}

function Get-CreatedSitePage {
    param(
        [string]$SiteId,
        [string]$PageId
    )

    $headers = New-GraphHeaders -MetadataNone
    $pageUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/pages/$PageId/microsoft.graph.sitePage"
    return Invoke-GraphRequestWithRetry -Uri $pageUrl -Headers $headers -Method GET
}

function Test-IsGraphNameConflict {
    param(
        [object]$ErrorRecord
    )

    if ($null -eq $ErrorRecord) {
        return $false
    }

    $statusCode = $null
    if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
        $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode
    }

    $message = [string]$ErrorRecord.Exception.Message
    $details = [string]$ErrorRecord.ErrorDetails.Message
    $combined = "$message`n$details"

    if ($statusCode -eq 409 -and ($combined -match 'nameAlreadyExists' -or $combined -match 'Name already exists')) {
        return $true
    }

    return $false
}

function New-RetryUniquePageName {
    param(
        [string]$CurrentName
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($CurrentName)
    $extension = [System.IO.Path]::GetExtension($CurrentName)
    if ([string]::IsNullOrWhiteSpace($extension)) {
        $extension = ".aspx"
    }

    $stamp = Get-Date -Format "yyyyMMddHHmmssfff"
    $randomPart = Get-Random -Minimum 100 -Maximum 999
    return "$baseName-r$stamp$randomPart$extension"
}

#endregion

#region Translation Flow

function Convert-SharePointPageToLanguage {
    param(
        [string]$SiteId,
        [object]$Page,
        [array]$ExistingPages
    )

    $startTime = Get-Date
    $webPartsProcessed = 0
    $textSegmentsTranslated = 0
    $textSegmentsRef = [ref]$textSegmentsTranslated
    $targetPageName = ""

    try {
        Write-Log "`nProcessing source page: $($Page.name)"
        Write-Log "Source page URL: $($Page.webUrl)"

        $sourceContent = Get-SitePageContent -SiteId $SiteId -PageId $Page.id
        
        # Extract the folder path from the source page URL FIRST, before naming
        $sourceFolderPath = Get-FolderPathFromWebUrl -WebUrl $Page.webUrl
        
        # Now get the translated page name, considering the folder path
        $targetPageName = Get-TranslatedPageName -SourceName $sourceContent.name -ExistingPages $ExistingPages -SourceFolderPath $sourceFolderPath
        if ($sourceFolderPath) {
            while (Test-SharePointFileExists -ServerRelativePath (Get-SitePagesFileServerRelativePath -FolderPath $sourceFolderPath -PageName $targetPageName)) {
                $newPageName = New-RetryUniquePageName -CurrentName $targetPageName
                Write-Log "Target page '$targetPageName' already exists in folder '$sourceFolderPath'. Retrying with '$newPageName'..." -Level WARNING
                $targetPageName = $newPageName
            }
        }
        Write-Log "Target page name: $targetPageName"

        if ($sourceContent.canvasLayout -and $sourceContent.canvasLayout.horizontalSections) {
            foreach ($section in $sourceContent.canvasLayout.horizontalSections) {
                foreach ($column in $section.columns) {
                    if ($column.webparts) {
                        $webPartsProcessed += @($column.webparts).Count
                    }
                }
            }
        }

        $createdPage = $null
        $targetWebUrl = ""
        $headers = New-GraphHeaders -MetadataNone
        $createUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/pages"
        $maxNameRetries = 5
        $nameRetryCount = 0

        if ($sourceFolderPath) {
            $createdPage = New-RestSubfolderTranslatedPage -Page $Page -SourceContent $sourceContent `
                -TargetPageName $targetPageName -TextSegmentsTranslated $textSegmentsRef -Publish:(-not $Draft)
            $targetWebUrl = $createdPage.webUrl
            $endTime = Get-Date
            Add-TranslationResult -SourcePage $Page.name -SourceWebUrl $Page.webUrl -TargetPage $targetPageName `
                -TargetWebUrl $targetWebUrl -Status "Success" -WebPartsProcessed $webPartsProcessed `
                -TextSegmentsTranslated $textSegmentsRef.Value -StartTime $startTime -EndTime $endTime

            Write-Log "Created translated REST subfolder page: $targetWebUrl" -Level SUCCESS
            return $createdPage
        }
        else {
            $payload = New-TranslatedPagePayload -SourcePageContent $sourceContent -NewPageName $targetPageName `
                -TextSegmentsTranslated $textSegmentsRef

            Write-Log "Creating translated SharePoint page with Graph..."
            while ($true) {
                try {
                    $body = $payload | ConvertTo-Json -Depth 100
                    $createdPage = Invoke-GraphRequestWithRetry -Uri $createUrl -Headers $headers -Method POST -Body $body
                    break
                }
                catch {
                    if ((Test-IsGraphNameConflict -ErrorRecord $_) -and $nameRetryCount -lt $maxNameRetries) {
                        $nameRetryCount++
                        $newPageName = New-RetryUniquePageName -CurrentName $targetPageName
                        Write-Log "Target page name '$targetPageName' already exists. Retrying with '$newPageName' ($nameRetryCount/$maxNameRetries)..." -Level WARNING
                        $targetPageName = $newPageName
                        $payload['name'] = $targetPageName
                        continue
                    }

                    throw
                }
            }

            $targetWebUrl = $createdPage.webUrl
        }

        if (-not $Draft) {
            Write-Log "Publishing translated page..."
            Publish-SitePage -SiteId $SiteId -PageId $createdPage.id
        }

        $verifiedPage = Get-CreatedSitePage -SiteId $SiteId -PageId $createdPage.id
        Write-Log "Verified created page ID: $($verifiedPage.id)"
        if ($verifiedPage.publishingState) {
            Write-Log "Created page publishing state: $($verifiedPage.publishingState.level) version $($verifiedPage.publishingState.versionId)"
        }

        if ($Draft) {
            Write-Log "The translated page was created as a draft. Draft pages created by app-only Graph calls may stay checked out and hidden from normal Site Pages views until published." -Level WARNING
        }

        $endTime = Get-Date
        Add-TranslationResult -SourcePage $Page.name -SourceWebUrl $Page.webUrl -TargetPage $targetPageName `
            -TargetWebUrl $targetWebUrl -Status "Success" -WebPartsProcessed $webPartsProcessed `
            -TextSegmentsTranslated $textSegmentsRef.Value -StartTime $startTime -EndTime $endTime

        Write-Log "Created translated page: $targetWebUrl" -Level SUCCESS
        return $createdPage
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Failed to translate page '$($Page.name)': $errorMsg" -Level ERROR
        if ($_.ErrorDetails.Message) {
            Write-Log "Error details: $($_.ErrorDetails.Message)" -Level ERROR
        }

        $endTime = Get-Date
        Add-TranslationResult -SourcePage $Page.name -SourceWebUrl $Page.webUrl -TargetPage $targetPageName `
            -Status "Failed" -ErrorMessage $errorMsg -WebPartsProcessed $webPartsProcessed `
            -TextSegmentsTranslated $textSegmentsRef.Value -StartTime $startTime -EndTime $endTime

        return $null
    }
}

function Start-PageTranslation {
    Assert-TranslatorConfigured
    if ($TranslatorAuthMode -eq 'Entra') {
        Get-CognitiveServicesToken
    }
    Test-TranslatorCredentials

    $siteId = Get-SiteIdFromUrl -Url $siteUrl
    $allPages = Get-AllSitePages -SiteId $siteId
    Write-Log "Found $($allPages.Count) ASPX page(s) in Site Pages."

    $pagesToTranslate = @(Resolve-PageReference -SiteId $siteId -Pages $allPages -Reference $PageName)
    Write-Log "Pages selected for translation: $($pagesToTranslate.Count)"

    $successCount = 0

    foreach ($page in $pagesToTranslate) {
        $createdPage = Convert-SharePointPageToLanguage -SiteId $siteId -Page $page -ExistingPages $allPages
        if ($createdPage) {
            $successCount++
            $allPages += $createdPage
        }
    }

    Write-Log "`n====== TRANSLATION SUMMARY ======" -Level SUCCESS
    Write-Log "Target language: $TargetLanguage" -Level INFO
    Write-Log "Pages selected: $($pagesToTranslate.Count)" -Level INFO
    Write-Log "Pages successfully created: $successCount" -Level SUCCESS
    Write-Log "Pages failed: $($pagesToTranslate.Count - $successCount)" -Level $(if ($pagesToTranslate.Count -eq $successCount) { "SUCCESS" } else { "WARNING" })
    Write-Log "=================================" -Level SUCCESS

    if ($global:translationResults.Count -gt 0) {
        $global:translationResults | Export-Csv -Path $csvLog -NoTypeInformation -Encoding UTF8
        Write-Log "Detailed results exported to: $csvLog" -Level SUCCESS
    }
}

#endregion

#region Main Execution

Write-Log "`n========================================" -Level INFO
Write-Log "ASPX Page Translation - Starting" -Level INFO
Write-Log "========================================" -Level INFO
Write-Log "Site URL: $siteUrl" -Level INFO
Write-Log "Input Page: $(if ($PageName) { $PageName } else { 'All ASPX pages' })" -Level INFO
Write-Log "Target Language: $TargetLanguage" -Level INFO
Write-Log "Output Name Suffix: $OutputNameSuffix" -Level INFO
Write-Log "App Auth Mode: $AppAuthMode" -Level INFO
Write-Log "Translator Auth Mode: $TranslatorAuthMode" -Level INFO
if ($TranslatorAuthMode -eq 'Entra') {
    Write-Log "Translator Token Tenant: $(Get-EffectiveTranslatorTenantId)" -Level INFO
}
Write-Log "Publish After Create: $(-not $Draft.IsPresent)" -Level INFO
Write-Log "Log File: $logFile" -Level INFO
Write-Log "CSV Report: $csvLog" -Level INFO
Write-Log "========================================`n" -Level INFO

Get-GraphToken
Start-PageTranslation

Write-Log "`n========================================" -Level SUCCESS
Write-Log "Translation Completed!" -Level SUCCESS
Write-Log "Log file saved to: $logFile" -Level SUCCESS
if ($global:translationResults.Count -gt 0) {
    Write-Log "CSV report saved to: $csvLog" -Level SUCCESS
}
Write-Log "========================================" -Level SUCCESS

#endregion
