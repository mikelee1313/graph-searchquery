<#
.SYNOPSIS
Searches Microsoft 365 (SharePoint + OneDrive) for files matching KQL query
ProgID:Media AND ProgID:Meeting and exports all hit metadata.

.DESCRIPTION
Uses Microsoft Graph Search API with app-only authentication.
The script pages through all results and exports each hit's full resource metadata
as JSON plus key columns to a CSV file.

.REQUIREMENTS
- Entra app registration with Microsoft Graph application permission: Files.Read.All (or broader)
- Admin consent granted for the app permission
- Certificate with private key available in CurrentUser or LocalMachine cert store
#>

[CmdletBinding()]
param(
    [Parameter()] [string] $TenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3',
    [Parameter()] [string] $ClientId = 'abc64618-283f-47ba-a185-50d935d51d57',
    [Parameter()] [string] $Thumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9',
    [Parameter()] [ValidateSet('CurrentUser', 'LocalMachine')] [string] $CertStore = 'LocalMachine',
    [Parameter()] [string] $SearchRegion = 'NAM',
    [Parameter()] [int]    $PageSize = 200,
    [Parameter()] [string] $KqlQuery = '(ProgID:Media AND ProgID:Meeting) OR FileType:mp4',
    [Parameter()] [switch] $EnableDriveItemEnrichment,
    [Parameter()] [switch] $IncludeExtendedColumns
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
    throw 'Thumbprint is empty. Pass -Thumbprint with a valid certificate thumbprint.'
}

if ($PageSize -lt 1 -or $PageSize -gt 1000) {
    throw 'PageSize must be between 1 and 1000.'
}

$date = Get-Date -Format 'yyyyMMddHHmmss'
$csvPath = Join-Path -Path $env:TEMP -ChildPath ("Search_Results_$date.csv")
$jsonPath = Join-Path -Path $env:TEMP -ChildPath ("Search_Results_$date.json")

function Get-GraphToken {
    param(
        [Parameter(Mandatory)] [string] $TenantId,
        [Parameter(Mandatory)] [string] $ClientId,
        [Parameter(Mandatory)] [string] $Thumbprint,
        [Parameter(Mandatory)] [string] $CertStore
    )

    function ConvertTo-Base64Url {
        param([Parameter(Mandatory)] [byte[]] $Bytes)
        return [Convert]::ToBase64String($Bytes).TrimEnd('=') -replace '\+', '-' -replace '/', '_'
    }

    function Get-CertificateByThumbprint {
        param(
            [Parameter(Mandatory)] [string] $Thumbprint,
            [Parameter(Mandatory)] [string] $Store
        )

        $normalized = ($Thumbprint -replace '\s', '').ToUpperInvariant()
        $path = ('Cert:\{0}\My\{1}' -f $Store, $normalized)
        if (-not (Test-Path $path)) {
            throw "Certificate '$normalized' not found in $path"
        }

        $cert = Get-Item -Path $path
        if (-not $cert.HasPrivateKey) {
            throw "Certificate '$normalized' does not have a private key."
        }
        return $cert
    }

    function New-ClientAssertion {
        param(
            [Parameter(Mandatory)] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate,
            [Parameter(Mandatory)] [string] $TenantId,
            [Parameter(Mandatory)] [string] $ClientId
        )

        $now = [DateTimeOffset]::UtcNow
        $aud = ('https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $TenantId)

        $x5t = ConvertTo-Base64Url -Bytes $Certificate.GetCertHash()
        $headerJson = @{
            alg = 'RS256'
            typ = 'JWT'
            x5t = $x5t
        } | ConvertTo-Json -Compress

        $payload = @{
            aud = $aud
            iss = $ClientId
            sub = $ClientId
            jti = ([Guid]::NewGuid()).Guid
            nbf = [int][Math]::Floor($now.ToUnixTimeSeconds())
            exp = [int][Math]::Floor($now.AddMinutes(10).ToUnixTimeSeconds())
        } | ConvertTo-Json -Compress

        $encodedHeader = ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($headerJson))
        $encodedPayload = ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($payload))
        $unsigned = ('{0}.{1}' -f $encodedHeader, $encodedPayload)

        $rsa = $null
        try {
            $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
        }
        catch {}

        # Fallback for older Windows PowerShell/.NET where extension invocation can fail.
        if ($null -eq $rsa -and $Certificate.PrivateKey -is [System.Security.Cryptography.RSA]) {
            $rsa = [System.Security.Cryptography.RSA]$Certificate.PrivateKey
        }

        if ($null -eq $rsa) {
            throw 'Unable to access RSA private key from certificate.'
        }

        $signatureBytes = $rsa.SignData(
            [Text.Encoding]::UTF8.GetBytes($unsigned),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
        )

        $encodedSignature = ConvertTo-Base64Url -Bytes $signatureBytes
        return ('{0}.{1}' -f $unsigned, $encodedSignature)
    }

    $tokenUri = ('https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $TenantId)
    $certificate = Get-CertificateByThumbprint -Thumbprint $Thumbprint -Store $CertStore
    $clientAssertion = New-ClientAssertion -Certificate $certificate -TenantId $TenantId -ClientId $ClientId

    $tokenBody = @{
        client_id             = $ClientId
        grant_type            = 'client_credentials'
        scope                 = 'https://graph.microsoft.com/.default'
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = $clientAssertion
    }

    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody -ContentType 'application/x-www-form-urlencoded'
    return $tokenResponse.access_token
}

function Invoke-GraphSearchPage {
    param(
        [Parameter(Mandatory)] [string] $AccessToken,
        [Parameter(Mandatory)] [string] $KqlQuery,
        [Parameter(Mandatory)] [int]    $From,
        [Parameter(Mandatory)] [int]    $Size,
        [Parameter(Mandatory)] [string] $SearchRegion
    )

    $uri = 'https://graph.microsoft.com/v1.0/search/query'
    $headers = @{ Authorization = "Bearer $AccessToken" }

    $payload = @{
        requests = @(
            @{
                entityTypes = @('driveItem')
                query = @{ queryString = $KqlQuery }
                fields = @('ProgID', 'ProgId', 'Path', 'Title', 'FileType', 'WebUrl')
                from = $From
                size = $Size
                region = $SearchRegion
                trimDuplicates = $false
                sharePointOneDriveOptions = @{ includeContent = 'sharedContent,privateContent' }
            }
        )
    } | ConvertTo-Json -Depth 10

    return Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $payload -ContentType 'application/json'
}

function Get-GraphDriveItemMetadata {
    param(
        [Parameter(Mandatory)] [string] $AccessToken,
        [Parameter(Mandatory)] [string] $DriveId,
        [Parameter(Mandatory)] [string] $ItemId
    )

    $uri = ('https://graph.microsoft.com/v1.0/drives/{0}/items/{1}?$expand=listItem($expand=fields)' -f $DriveId, $ItemId)
    $headers = @{ Authorization = "Bearer $AccessToken" }
    return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
}

function Get-GraphDriveItemMetadataByWebUrl {
    param(
        [Parameter(Mandatory)] [string] $AccessToken,
        [Parameter(Mandatory)] [string] $WebUrl
    )

    # Graph shares API expects a URL-safe base64 token prefixed with u!
    $encoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($WebUrl)).TrimEnd('=') -replace '\+', '-' -replace '/', '_'
    $shareId = 'u!' + $encoded
    $uri = ('https://graph.microsoft.com/v1.0/shares/{0}/driveItem?$expand=listItem($expand=fields)' -f $shareId)
    $headers = @{ Authorization = "Bearer $AccessToken" }
    return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
}

function Get-AnyValue {
    param(
        [Parameter()] $Object,
        [Parameter(Mandatory)] [string] $Name
    )

    if ($null -eq $Object) { return $null }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }

        foreach ($k in $Object.Keys) {
            if ([string]::Equals([string]$k, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $Object[$k]
            }
        }

        return $null
    }

    $prop = $Object.PSObject.Properties | Where-Object {
        [string]::Equals($_.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)
    } | Select-Object -First 1

    if ($null -ne $prop) { return $prop.Value }
    return $null
}

function Get-NestedValue {
    param(
        [Parameter()] $Object,
        [Parameter(Mandatory)] [string[]] $Path
    )

    $current = $Object
    foreach ($segment in $Path) {
        $current = Get-AnyValue -Object $current -Name $segment
        if ($null -eq $current) { return $null }
    }
    return $current
}

function Get-FirstNonEmpty {
    param(
        [Parameter()] [object[]] $Values
    )

    foreach ($value in $Values) {
        if ($null -eq $value) { continue }

        $text = [string]$value
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            return $value
        }
    }

    return $null
}

function Get-BestFileName {
    param(
        [Parameter()] $Resource,
        [Parameter()] $HitFields,
        [Parameter()] $SearchListItemFields
    )

    $candidate = Get-FirstNonEmpty -Values @(
        (Get-AnyValue -Object $Resource -Name 'name'),
        (Get-AnyValue -Object $HitFields -Name 'fileName'),
        (Get-AnyValue -Object $SearchListItemFields -Name 'FileLeafRef'),
        (Get-AnyValue -Object $SearchListItemFields -Name 'LinkFilename')
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$candidate)) {
        return [string]$candidate
    }

    $title = Get-FirstNonEmpty -Values @(
        (Get-AnyValue -Object $HitFields -Name 'title'),
        (Get-AnyValue -Object $SearchListItemFields -Name 'title')
    )
    $fileType = Get-FirstNonEmpty -Values @(
        (Get-AnyValue -Object $HitFields -Name 'fileType'),
        (Get-AnyValue -Object $SearchListItemFields -Name 'fileType')
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$title) -and -not [string]::IsNullOrWhiteSpace([string]$fileType)) {
        $titleText = [string]$title
        $ext = [string]$fileType
        if ($titleText -notmatch ("\.{0}$" -f [Regex]::Escape($ext))) {
            return ('{0}.{1}' -f $titleText, $ext)
        }

        return $titleText
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$title)) {
        return [string]$title
    }

    return ''
}

Write-Host "Acquiring Graph token..." -ForegroundColor Cyan
$accessToken = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -Thumbprint $Thumbprint -CertStore $CertStore
Write-Host 'Token acquired.' -ForegroundColor Green

Write-Host "Running KQL query: $KqlQuery" -ForegroundColor Cyan

$allHits = @()
$from = 0
$loopCount = 0

while ($true) {
    $loopCount++
    $response = Invoke-GraphSearchPage -AccessToken $accessToken -KqlQuery $KqlQuery -From $from -Size $PageSize -SearchRegion $SearchRegion

    # Some environments return Graph JSON as a string; parse it so property checks work.
    if ($response -is [string]) {
        try {
            try {
                # PowerShell 7+
                $response = $response | ConvertFrom-Json -Depth 30 -ErrorAction Stop
            }
            catch {
                try {
                    # Windows PowerShell 5.1 fallback (no -Depth parameter)
                    $response = $response | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    # Handles duplicate keys that differ only by case (for example progID/progId).
                    $response = $response | ConvertFrom-Json -AsHashtable -ErrorAction Stop
                }
            }
        }
        catch {
            throw ('Search response was a string but could not be parsed as JSON. Raw response: {0}' -f $response)
        }
    }

    $responseValue = $null
    if ($response -is [System.Array]) {
        $responseValue = $response
    }
    else {
        $responseValue = @(Get-AnyValue -Object $response -Name 'value')
        if ($responseValue.Count -eq 1 -and $null -eq $responseValue[0]) {
            $responseValue = @(Get-AnyValue -Object $response -Name 'Value')
        }
    }

    if ($responseValue.Count -eq 1 -and $responseValue[0] -is [System.Array]) {
        $responseValue = @($responseValue[0])
    }

    if (-not $responseValue -or ($responseValue.Count -eq 1 -and $null -eq $responseValue[0])) {
        $raw = ($response | ConvertTo-Json -Depth 20 -Compress)
        throw ('Search response did not contain a top-level value property. Raw response: {0}' -f $raw)
    }

    if ($responseValue.Count -eq 0) {
        break
    }

    $firstResponseItem = $responseValue[0]
    $hitsContainers = Get-AnyValue -Object $firstResponseItem -Name 'hitsContainers'
    if ($null -eq $firstResponseItem -or $null -eq $hitsContainers) {
        $rawItem = ($firstResponseItem | ConvertTo-Json -Depth 20 -Compress)
        throw ('Search response item did not contain hitsContainers. Raw item: {0}' -f $rawItem)
    }

    $container = @($hitsContainers)[0]
    $hits = @(Get-AnyValue -Object $container -Name 'hits')
    if ($hits.Count -eq 1 -and $hits[0] -is [System.Array]) {
        $hits = @($hits[0])
    }

    if ($hits.Count -eq 0) {
        break
    }

    $allHits += $hits
    Write-Host ("  Page {0}: retrieved {1} hit(s), total so far: {2}" -f $loopCount, $hits.Count, $allHits.Count) -ForegroundColor DarkGray

    $moreResultsAvailable = Get-AnyValue -Object $container -Name 'moreResultsAvailable'
    if (-not $moreResultsAvailable) {
        break
    }

    $from += $PageSize
}

if ($allHits.Count -eq 0) {
    Write-Host 'No matching files found.' -ForegroundColor Yellow
    return
}

# Export key columns plus full search hit and full drive item metadata.
$csvRows = foreach ($hit in $allHits) {
    $resource = Get-AnyValue -Object $hit -Name 'resource'
    $hitFields = Get-AnyValue -Object $hit -Name 'fields'
    $searchListItemFields = Get-NestedValue -Object $resource -Path @('listItem', 'fields')

    $itemMetadata = $null
    $metadataError = $null
    $driveId = Get-NestedValue -Object $resource -Path @('parentReference', 'driveId')
    $itemId = Get-AnyValue -Object $resource -Name 'id'
    $resourceWebUrl = Get-AnyValue -Object $resource -Name 'webUrl'

    if ($EnableDriveItemEnrichment -and -not [string]::IsNullOrWhiteSpace($driveId) -and -not [string]::IsNullOrWhiteSpace($itemId)) {
        try {
            $itemMetadata = Get-GraphDriveItemMetadata -AccessToken $accessToken -DriveId $driveId -ItemId $itemId
        }
        catch {
            $metadataError = $_.Exception.Message
            if ([string]::IsNullOrWhiteSpace($metadataError) -and $_.ErrorDetails -and $_.ErrorDetails.Message) {
                $metadataError = $_.ErrorDetails.Message
            }
        }
    }
    elseif ($EnableDriveItemEnrichment -and -not [string]::IsNullOrWhiteSpace($resourceWebUrl)) {
        try {
            $itemMetadata = Get-GraphDriveItemMetadataByWebUrl -AccessToken $accessToken -WebUrl $resourceWebUrl
        }
        catch {
            $metadataError = $_.Exception.Message
            if ([string]::IsNullOrWhiteSpace($metadataError) -and $_.ErrorDetails -and $_.ErrorDetails.Message) {
                $metadataError = $_.ErrorDetails.Message
            }
        }
    }
    elseif ($EnableDriveItemEnrichment) {
        $metadataError = 'No driveId/itemId or webUrl available in hit resource for metadata lookup.'
    }

    $searchHitJson = ($hit | ConvertTo-Json -Depth 25 -Compress)
    $resourceJson = ($resource | ConvertTo-Json -Depth 25 -Compress)
    $itemMetadataJson = if ($null -ne $itemMetadata) { $itemMetadata | ConvertTo-Json -Depth 25 -Compress } else { '' }
    $itemMetadataFields = Get-NestedValue -Object $itemMetadata -Path @('listItem', 'fields')
    $listItemFieldsJson = if ($null -ne $itemMetadata -and $null -ne $itemMetadata.listItem -and $null -ne $itemMetadata.listItem.fields) {
        $itemMetadata.listItem.fields | ConvertTo-Json -Depth 25 -Compress
    }
    elseif ($null -ne $searchListItemFields) {
        $searchListItemFields | ConvertTo-Json -Depth 25 -Compress
    }
    else { '' }

    $isForbiddenMetadataCall = -not [string]::IsNullOrWhiteSpace($metadataError) -and $metadataError -match '403|Forbidden'

    $progIdRaw = Get-AnyValue -Object $hitFields -Name 'progID'
    if ([string]::IsNullOrWhiteSpace([string]$progIdRaw)) {
        $progIdRaw = Get-AnyValue -Object $searchListItemFields -Name 'progID'
    }
    if ([string]::IsNullOrWhiteSpace([string]$progIdRaw)) {
        $progIdRaw = Get-AnyValue -Object $itemMetadataFields -Name 'progID'
    }

    $progIdText = if ($null -ne $progIdRaw) { [string]$progIdRaw } else { '' }
    $hasProgIdMedia = $progIdText -match '(?i)\bMedia\b'
    $hasProgIdMeeting = $progIdText -match '(?i)\bMeeting\b'
    $isMissingProgIdMetadata = -not ($hasProgIdMedia -and $hasProgIdMeeting)

    $resolvedFileName = Get-BestFileName -Resource $resource -HitFields $hitFields -SearchListItemFields $searchListItemFields
    $resolvedWebUrl = Get-FirstNonEmpty -Values @(
        (Get-AnyValue -Object $resource -Name 'webUrl'),
        (Get-AnyValue -Object $hitFields -Name 'webUrl'),
        (Get-AnyValue -Object $searchListItemFields -Name 'path')
    )
    $resolvedPath = Get-FirstNonEmpty -Values @(
        (Get-AnyValue -Object $hitFields -Name 'path'),
        (Get-AnyValue -Object $searchListItemFields -Name 'path'),
        $resolvedWebUrl
    )
    $resolvedFileType = Get-FirstNonEmpty -Values @(
        (Get-NestedValue -Object $resource -Path @('file', 'mimeType')),
        (Get-AnyValue -Object $hitFields -Name 'fileType'),
        (Get-AnyValue -Object $searchListItemFields -Name 'fileType')
    )

    # Print all metadata JSON for each result in console output.
    Write-Host ("`n========== Metadata for: {0} ==========" -f $resourceWebUrl) -ForegroundColor Cyan
    Write-Host ('Search hit metadata: {0}' -f $searchHitJson)
    if ($itemMetadataJson) {
        Write-Host ('Drive item metadata: {0}' -f $itemMetadataJson)
    }
    elseif ($EnableDriveItemEnrichment) {
        if ([string]::IsNullOrWhiteSpace($metadataError)) {
            $metadataError = 'Metadata call returned no content.'
        }

        if ($isForbiddenMetadataCall) {
            Write-Host ('Info: additional driveItem API metadata is blocked (403). Using search/listItem metadata for this file.') -ForegroundColor DarkYellow
        }
        else {
            Write-Warning ('Drive item metadata could not be retrieved: {0}' -f $metadataError)
        }
    }

    if ($isMissingProgIdMetadata) {
        if (-not $hasProgIdMedia -and -not $hasProgIdMeeting) {
            Write-Host 'Flag: missing "ProgID:Media" and "ProgID:Meeting" metadata.' -ForegroundColor Red
        }
        elseif (-not $hasProgIdMedia) {
            Write-Host 'Flag: missing "ProgID:Media" metadata.' -ForegroundColor Red
        }
        else {
            Write-Host 'Flag: missing "ProgID:Meeting" metadata.' -ForegroundColor Red
        }
    }

    $row = [ordered]@{
        FileName             = $resolvedFileName
        WebUrl               = $resolvedWebUrl
        HitPath              = $resolvedPath
        FileMimeType         = $resolvedFileType
        HitProgID            = $progIdText
        HasProgIDMedia       = $hasProgIdMedia
        HasProgIDMeeting     = $hasProgIdMeeting
        MissingProgIDMetadata = $isMissingProgIdMetadata
        Rank                 = (Get-AnyValue -Object $hit -Name 'rank')
        HitId                = (Get-AnyValue -Object $hit -Name 'hitId')
        Summary              = (Get-AnyValue -Object $hit -Name 'summary')
        MetadataSource       = if ($itemMetadataJson) { 'DriveItemApi' } else { 'SearchHit' }
        QueryUsed            = $KqlQuery
    }

    if ($IncludeExtendedColumns) {
        $row['Id'] = (Get-AnyValue -Object $resource -Name 'id')
        $row['DriveId'] = (Get-NestedValue -Object $resource -Path @('parentReference', 'driveId'))
        $row['SiteId'] = (Get-NestedValue -Object $resource -Path @('parentReference', 'siteId'))
        $row['CreatedDateUtc'] = (Get-AnyValue -Object $resource -Name 'createdDateTime')
        $row['ModifiedDateUtc'] = (Get-AnyValue -Object $resource -Name 'lastModifiedDateTime')
        $row['CreatedBy'] = (Get-NestedValue -Object $resource -Path @('createdBy', 'user', 'displayName'))
        $row['ModifiedBy'] = (Get-NestedValue -Object $resource -Path @('lastModifiedBy', 'user', 'displayName'))
        $row['Size'] = (Get-AnyValue -Object $resource -Name 'size')
        $row['ResourceJson'] = $resourceJson
        $row['ListItemFieldsJson'] = $listItemFieldsJson
        $row['FullDriveItemMetadataJson'] = $itemMetadataJson
        $row['SearchHitJson'] = $searchHitJson
        $row['MetadataReadError'] = if (-not $EnableDriveItemEnrichment -or $isForbiddenMetadataCall) { '' } else { $metadataError }
    }

    [PSCustomObject]$row
}

$csvRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
$allHits | ConvertTo-Json -Depth 25 | Set-Content -Path $jsonPath -Encoding UTF8

Write-Host "Search finished successfully." -ForegroundColor Green
Write-Host "Total matching files: $($allHits.Count)" -ForegroundColor Green
Write-Host "CSV output: $csvPath" -ForegroundColor Yellow
Write-Host "JSON output: $jsonPath" -ForegroundColor Yellow
