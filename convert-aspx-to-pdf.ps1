<#
.SYNOPSIS
    Converts SharePoint ASPX pages to PDF using Microsoft Graph API.

.DESCRIPTION
    This script retrieves ASPX pages from a SharePoint site's SitePages library and converts them to PDF format.
    It uses Graph API to extract page content including text, images, and link web parts, embedding all content
    as base64 data to ensure proper rendering in the PDF output.

.PARAMETER tenantId
    The Azure AD tenant ID for authentication.

.PARAMETER clientId
    The client ID of the Azure AD application.

.PARAMETER Thumbprint
    The certificate thumbprint for authentication.

.PARAMETER CertStore
    Certificate store location: "LocalMachine" or "CurrentUser" (default: "LocalMachine").

.PARAMETER siteUrl
    The SharePoint site URL to convert pages from.

.PARAMETER outputFolder
    Local folder path where converted PDFs will be saved (default: user's temp directory).

.EXAMPLE
    .\convert-aspx-to-pdf.ps1
    Converts all ASPX pages from the configured SharePoint site to PDF.

.NOTES
    Author: Mike Lee
    Date: November 21, 2025
    Required Permissions: Sites.Read.All, Files.Read.All
#>

#region Configuration
##############################################################
#                  CONFIGURATION SECTION                     #
##############################################################

$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'
$clientId = 'abc64618-283f-47ba-a185-50d935d51d57'
$Thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"
$CertStore = 'LocalMachine'
$siteUrl = "https://m365cpi13246019.sharepoint.com/sites/SalesandMarketing"
$outputFolder = $env:TEMP
$logFolder = Join-Path -Path $env:TEMP -ChildPath "ASPX_Conversion_Logs"  # Centralized log location

##############################################################
#endregion

#region Initialization

$date = Get-Date -Format "yyyyMMddHHmmss"
$convertedOutputDir = Join-Path -Path $outputFolder -ChildPath ("ConvertedASPX_" + $date)
if (-not (Test-Path $convertedOutputDir)) {
    New-Item -ItemType Directory -Path $convertedOutputDir -Force | Out-Null
}

# Create log folder if it doesn't exist
if (-not (Test-Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
}

# Log files
$logFile = Join-Path -Path $logFolder -ChildPath "conversion_log_$date.log"
$csvLog = Join-Path -Path $logFolder -ChildPath "conversion_summary_$date.csv"

$global:token = ""
$global:conversionResults = @()

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
    
    # Write to console with color
    switch ($Level) {
        'INFO' { Write-Host $logEntry -ForegroundColor Cyan }
        'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
        'ERROR' { Write-Host $logEntry -ForegroundColor Red }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
    }
    
    # Write to log file
    Add-Content -Path $logFile -Value $logEntry
}

function Add-ConversionResult {
    param(
        [string]$FileName,
        [string]$WebUrl,
        [string]$Status,
        [string]$OutputPath = "",
        [string]$ErrorMessage = "",
        [int]$WebPartsProcessed = 0,
        [int]$ImagesEmbedded = 0,
        [int]$LinksProcessed = 0,
        [datetime]$StartTime,
        [datetime]$EndTime
    )
    
    $result = [PSCustomObject]@{
        Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        FileName          = $FileName
        WebUrl            = $WebUrl
        Status            = $Status
        OutputPath        = $OutputPath
        ErrorMessage      = $ErrorMessage
        WebPartsProcessed = $WebPartsProcessed
        ImagesEmbedded    = $ImagesEmbedded
        LinksProcessed    = $LinksProcessed
        ProcessingTime    = ($EndTime - $StartTime).TotalSeconds
        StartTime         = $StartTime.ToString("yyyy-MM-dd HH:mm:ss")
        EndTime           = $EndTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
    
    $global:conversionResults += $result
}

#endregion

#region Authentication

function Get-GraphToken {
    Write-Log "Connecting to Microsoft Graph using Certificate authentication..."
    
    $certPath = "Cert:\$CertStore\My\$Thumbprint"
    if (-not (Test-Path $certPath)) {
        Write-Log "Certificate not found at $certPath" -Level ERROR
        exit 1
    }
    
    $cert = Get-Item $certPath
    Write-Log "Certificate found in $CertStore\My store" -Level SUCCESS
    
    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    
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
        scope                 = "https://graph.microsoft.com/.default"
        grant_type            = "client_credentials"
    }
    
    $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
    $global:token = $response.access_token
    
    Write-Log "Successfully connected using Certificate authentication" -Level SUCCESS
}

#endregion

#region Throttling Helpers

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
            
            if ($Body) {
                $params['Body'] = $Body
            }
            
            if ($Method -eq "GET" -and $Uri -match '/content') {
                # For binary content, use Invoke-WebRequest
                return Invoke-WebRequest @params
            }
            else {
                return Invoke-RestMethod @params
            }
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            
            if ($statusCode -eq 429) {
                # Throttled - check for Retry-After header
                $retryAfter = $null
                
                if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers['Retry-After']) {
                    $retryAfter = [int]$_.Exception.Response.Headers['Retry-After']
                    Write-Host -ForegroundColor Yellow "  Throttled (429). Retry-After: $retryAfter seconds"
                }
                else {
                    # Exponential backoff if no Retry-After header
                    $retryAfter = [Math]::Pow(2, $retryCount) * $baseDelay
                    Write-Host -ForegroundColor Yellow "  Throttled (429). Using exponential backoff: $retryAfter seconds"
                }
                
                if ($retryCount -lt $MaxRetries) {
                    Write-Host -ForegroundColor Cyan "  Waiting $retryAfter seconds before retry ($($retryCount + 1)/$MaxRetries)..."
                    Start-Sleep -Seconds $retryAfter
                    $retryCount++
                    continue
                }
                else {
                    Write-Host -ForegroundColor Red "  Max retries reached. Request failed."
                    throw
                }
            }
            elseif ($statusCode -ge 500 -and $statusCode -lt 600) {
                # Server error - retry with exponential backoff
                $retryAfter = [Math]::Pow(2, $retryCount) * $baseDelay
                Write-Host -ForegroundColor Yellow "  Server error ($statusCode). Retrying in $retryAfter seconds..."
                
                if ($retryCount -lt $MaxRetries) {
                    Start-Sleep -Seconds $retryAfter
                    $retryCount++
                    continue
                }
                else {
                    Write-Host -ForegroundColor Red "  Max retries reached. Request failed."
                    throw
                }
            }
            else {
                # Other error - don't retry
                throw
            }
        }
    }
}

#endregion

#region ASPX to PDF Conversion

function ConvertAspxToPdf {
    param(
        [string]$fileName,
        [string]$webUrl,
        [object]$fileResource
    )
    
    $startTime = Get-Date
    $webPartsCount = 0
    $imagesCount = 0
    $linksCount = 0
    
    Write-Log "Processing ASPX file: $fileName"
    Write-Log "ASPX Web URL: $webUrl"
    
    $headers = @{
        "Authorization" = "Bearer $global:token"
        "Content-Type"  = "application/json"
    }
    
    $siteId = $null
    $parsedUrl = [System.Uri]$webUrl
    $hostname = $parsedUrl.Host
    $sitePath = $parsedUrl.AbsolutePath -replace '/SitePages/.*$', ''
    
    Write-Log "Attempting to retrieve page content via Graph API..."
    $siteIdUrl = "https://graph.microsoft.com/v1.0/sites/$hostname`:$sitePath"
    $siteResponse = Invoke-GraphRequestWithRetry -Uri $siteIdUrl -Headers $headers -Method GET
    $siteId = $siteResponse.id
    Write-Log "Site ID: $siteId"
    
    $pagesUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/pages"
    $pagesResponse = Invoke-GraphRequestWithRetry -Uri $pagesUrl -Headers $headers -Method GET
    
    $pageName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    Write-Log "Looking for page with name: $pageName"
    
    # Find the page - try different properties
    $page = $pagesResponse.value | Where-Object { 
        $_.name -eq $pageName -or 
        $_.name -eq $fileName -or
        $_.title -eq $pageName
    }
    
    if (-not $page -and $pagesResponse.value.Count -gt 0) {
        Write-Log "Available pages:" -Level WARNING
        $pagesResponse.value | ForEach-Object {
            Write-Log "  Name: '$($_.name)', Title: '$($_.title)', ID: $($_.id)" -Level WARNING
        }
        # Try case-insensitive match
        $page = $pagesResponse.value | Where-Object { 
            $_.name -like $pageName -or 
            $_.name -like $fileName -or
            $_.title -like $pageName
        }
    }
    
    if ($page) {
        Write-Log "Found page via Graph API (Name: '$($page.name)', Title: '$($page.title)')" -Level SUCCESS
        
        # Extract just the first GUID from the page ID (Graph API sometimes returns an array of GUIDs)
        $pageId = $page.id
        
        # If it's an array, take the first element
        if ($pageId -is [array]) {
            Write-Log "Raw page ID from API: $($pageId -join ' ')"
            $pageId = $pageId[0]
            Write-Log "Multiple GUIDs detected, using first: $pageId"
        }
        else {
            Write-Log "Raw page ID from API: $pageId"
        }
        
        Write-Log "Using page ID: $pageId"
        
        try {
            Write-Log "Trying beta API for full content..."
            $webPartsUrl = "https://graph.microsoft.com/beta/sites/$siteId/pages/$pageId/microsoft.graph.sitePage/webParts"
            $webPartsResponse = Invoke-GraphRequestWithRetry -Uri $webPartsUrl -Headers $headers -Method GET
            Write-Log "Successfully retrieved web parts from beta API" -Level SUCCESS
            
            $htmlContent = "<html><head><meta charset='UTF-8'><title>$fileName</title><style>body{font-family:Segoe UI,Arial,sans-serif;margin:20px;line-height:1.6;}h1,h2,h3{color:#333;}a{color:#0078d4;}</style></head><body>"
            $htmlContent += "<h1>$($page.title)</h1>"
            
            foreach ($webPart in $webPartsResponse.value) {
                $webPartsCount++
                Write-Log "  Processing web part: $($webPart.id)"
                
                if ($webPart.innerHtml) {
                    Write-Log "    Adding innerHtml ($($webPart.innerHtml.Length) chars)"
                    $htmlContent += $webPart.innerHtml
                }
                
                if ($webPart.data) {
                    Write-Log "    Has data field, attempting to parse..."
                    try {
                        $data = $webPart.data
                        
                        # Image web part
                        if ($data.properties.uniqueId) {
                            $uniqueId = $data.properties.uniqueId
                            $listId = $data.properties.listId
                            $imgUrl = "https://$hostname/_layouts/15/download.aspx?UniqueId=$uniqueId"
                            Write-Log "      Found image URL: $imgUrl"
                            
                            try {
                                $downloadUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items/$uniqueId/driveItem/content"
                                Write-Log "      Downloading image via Graph API..."
                                
                                $imageResponse = Invoke-GraphRequestWithRetry -Uri $downloadUrl -Headers $headers -Method GET
                                
                                if ($imageResponse.Content -and $imageResponse.Content.Length -gt 0) {
                                    $imageBytes = $imageResponse.Content
                                    
                                    $contentType = "image/jpeg"
                                    if ($imageResponse.Headers.'Content-Type') {
                                        $contentType = $imageResponse.Headers.'Content-Type'
                                    }
                                    
                                    $base64 = [Convert]::ToBase64String($imageBytes)
                                    $imgSrc = "data:$contentType;base64,$base64"
                                    
                                    Write-Log "      Image downloaded and embedded as base64 ($($imageBytes.Length) bytes, type: $contentType)" -Level SUCCESS
                                    $htmlContent += "<div style='margin: 20px 0; text-align: center;'><img src='$imgSrc' style='max-width:100%; height:auto;' alt='Image' /></div>"
                                    $imagesCount++
                                }
                            }
                            catch {
                                Write-Log "      Could not download image via Graph API: $($_.Exception.Message)" -Level WARNING
                            }
                        }
                        
                        # Link web part
                        if ($data.serverProcessedContent) {
                            $linkUrl = $null
                            $linkImageUrl = $null
                            $linkTitle = $data.properties.title
                            $linkDescription = $data.properties.description
                            
                            if ($data.serverProcessedContent.links) {
                                $linksObj = $data.serverProcessedContent.links
                                if ($linksObj.key -eq 'url' -and $linksObj.value) {
                                    $linkUrl = $linksObj.value
                                    Write-Log "      Found link URL: $linkUrl" -Level SUCCESS
                                }
                            }
                            
                            if ($data.serverProcessedContent.imageSources) {
                                $imgSourcesObj = $data.serverProcessedContent.imageSources
                                if ($imgSourcesObj.key -eq 'imageURL' -and $imgSourcesObj.value) {
                                    $linkImageUrl = $imgSourcesObj.value
                                    Write-Log "      Found link preview image: $linkImageUrl" -Level SUCCESS
                                }
                            }
                            
                            if ($linkUrl) {
                                $linksCount++
                                $htmlContent += "<div style='border: 1px solid #e1e1e1; border-radius: 4px; padding: 16px; margin: 20px 0; background: #fafafa;'>"
                                
                                if ($linkImageUrl) {
                                    try {
                                        Write-Log "      Downloading link preview image..."
                                        $imgResponse = Invoke-WebRequest -Uri $linkImageUrl -UseBasicParsing
                                        $imgBytes = $imgResponse.Content
                                        
                                        $imgContentType = "image/jpeg"
                                        if ($imgResponse.Headers.'Content-Type') {
                                            $imgContentType = $imgResponse.Headers.'Content-Type'
                                        }
                                        elseif ($linkImageUrl -match '\.(jpg|jpeg)$') { $imgContentType = "image/jpeg" }
                                        elseif ($linkImageUrl -match '\.png$') { $imgContentType = "image/png" }
                                        elseif ($linkImageUrl -match '\.gif$') { $imgContentType = "image/gif" }
                                        
                                        $imgBase64 = [Convert]::ToBase64String($imgBytes)
                                        $imgDataUrl = "data:$imgContentType;base64,$imgBase64"
                                        
                                        Write-Log "      Link image embedded as base64 ($($imgBytes.Length) bytes, type: $imgContentType)" -Level SUCCESS
                                        $htmlContent += "<div style='margin-bottom: 12px;'><img src='$imgDataUrl' style='max-width: 100%; height: auto; border-radius: 4px;' /></div>"
                                    }
                                    catch {
                                        Write-Log "      Could not download link preview image: $($_.Exception.Message)" -Level WARNING
                                        $htmlContent += "<div style='margin-bottom: 12px;'><img src='$linkImageUrl' style='max-width: 100%; height: auto; border-radius: 4px;' /></div>"
                                    }
                                }
                                
                                if ($linkTitle) {
                                    $htmlContent += "<div style='font-weight: 600; margin-bottom: 8px;'>$linkTitle</div>"
                                }
                                
                                if ($linkDescription) {
                                    $htmlContent += "<div style='color: #666; margin-bottom: 8px; font-size: 14px;'>$linkDescription</div>"
                                }
                                
                                $htmlContent += "<div><a href='$linkUrl' style='color: #0078d4; text-decoration: none; font-size: 14px;'>$linkUrl</a></div>"
                                $htmlContent += "</div>"
                                
                            }
                        }
                        
                        if ($data.innerHTML) {
                            $htmlContent += $data.innerHTML
                        }
                    }
                    catch {
                        Write-Log "    Failed to parse data: $($_.Exception.Message)" -Level WARNING
                    }
                }
            }
            
            $htmlContent += "</body></html>"
            
            # Save HTML file
            $htmlFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + ".html"
            $htmlFilePath = Join-Path -Path $convertedOutputDir -ChildPath $htmlFileName
            $htmlContent | Out-File -FilePath $htmlFilePath -Encoding UTF8
            
            Write-Log "HTML file created: $htmlFileName" -Level SUCCESS
            
            # Convert to PDF using Graph API
            $convertedFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + ".pdf"
            $convertedFilePath = Join-Path -Path $convertedOutputDir -ChildPath $convertedFileName
            
            try {
                Write-Log "Converting HTML to PDF using Graph API..."
                
                # Get drive ID
                $driveUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/drive"
                $driveInfo = Invoke-GraphRequestWithRetry -Uri $driveUrl -Headers $headers -Method GET
                $driveId = $driveInfo.id
                
                # Upload HTML to temp folder
                $tempFolderName = "TempConversion"
                $uploadUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/$tempFolderName/$($htmlFileName):/content"
                $htmlBytes = [System.Text.Encoding]::UTF8.GetBytes($htmlContent)
                
                $uploadResponse = Invoke-GraphRequestWithRetry -Uri $uploadUrl -Method PUT -Headers $headers -Body $htmlBytes -ContentType "text/html"
                $driveItemId = $uploadResponse.id
                
                # Convert to PDF
                $convertUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$driveItemId/content?format=pdf"
                $pdfResponse = Invoke-GraphRequestWithRetry -Uri $convertUrl -Method GET -Headers $headers
                
                # Save PDF
                [System.IO.File]::WriteAllBytes($convertedFilePath, $pdfResponse.Content)
                
                # Clean up temp file
                $deleteUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$driveItemId"
                Invoke-GraphRequestWithRetry -Uri $deleteUrl -Method DELETE -Headers $headers | Out-Null
                
                Write-Log "Successfully converted $fileName to PDF" -Level SUCCESS
                Write-Log "PDF saved: $convertedFilePath" -Level SUCCESS
                
                $endTime = Get-Date
                Add-ConversionResult -FileName $fileName -WebUrl $webUrl -Status "Success" -OutputPath $convertedFilePath `
                    -WebPartsProcessed $webPartsCount -ImagesEmbedded $imagesCount -LinksProcessed $linksCount `
                    -StartTime $startTime -EndTime $endTime
                
                return $convertedFilePath
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Graph API conversion failed: $errorMsg" -Level ERROR
                if ($_.ErrorDetails.Message) {
                    Write-Log "Error details: $($_.ErrorDetails.Message)" -Level ERROR
                }
                
                $endTime = Get-Date
                Add-ConversionResult -FileName $fileName -WebUrl $webUrl -Status "Failed" -ErrorMessage $errorMsg `
                    -WebPartsProcessed $webPartsCount -ImagesEmbedded $imagesCount -LinksProcessed $linksCount `
                    -StartTime $startTime -EndTime $endTime
                
                return $null
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Beta API failed: $errorMsg" -Level ERROR
            
            $endTime = Get-Date
            Add-ConversionResult -FileName $fileName -WebUrl $webUrl -Status "Failed" -ErrorMessage $errorMsg `
                -WebPartsProcessed $webPartsCount -ImagesEmbedded $imagesCount -LinksProcessed $linksCount `
                -StartTime $startTime -EndTime $endTime
            
            return $null
        }
    }
    else {
        $errorMsg = "Page not found via Graph API"
        Write-Log $errorMsg -Level ERROR
        
        $endTime = Get-Date
        Add-ConversionResult -FileName $fileName -WebUrl $webUrl -Status "Failed" -ErrorMessage $errorMsg `
            -StartTime $startTime -EndTime $endTime
        
        return $null
    }
}

#endregion

#region Get SitePages

function GetSitePagesDirectly {
    $headers = @{
        "Authorization" = "Bearer $global:token"
        "Content-Type"  = "application/json"
    }
    
    Write-Log "Getting ASPX files directly from SitePages library..."
    
    $parsedUrl = [System.Uri]$siteUrl
    $hostname = $parsedUrl.Host
    $sitePath = $parsedUrl.AbsolutePath
    
    Write-Log "Getting site ID for: $hostname`:$sitePath"
    $siteIdUrl = "https://graph.microsoft.com/v1.0/sites/$hostname`:$sitePath"
    $siteResponse = Invoke-GraphRequestWithRetry -Uri $siteIdUrl -Headers $headers -Method GET
    $siteId = $siteResponse.id
    Write-Log "Site ID: $siteId"
    
    Write-Log "Getting SitePages library..."
    $listsUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/lists?`$filter=displayName eq 'Site Pages'"
    $listsResponse = Invoke-GraphRequestWithRetry -Uri $listsUrl -Headers $headers -Method GET
    
    if ($listsResponse.value.Count -eq 0) {
        Write-Log "SitePages library not found" -Level ERROR
        return
    }
    
    $sitePagesListId = $listsResponse.value[0].id
    Write-Log "SitePages library ID: $sitePagesListId"
    
    Write-Log "Getting items from SitePages library..."
    $itemsUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$sitePagesListId/items?`$expand=fields"
    $itemsResponse = Invoke-GraphRequestWithRetry -Uri $itemsUrl -Headers $headers -Method GET
    
    Write-Log "Found $($itemsResponse.value.Count) items in SitePages"
    
    $fileCount = 0
    $convertedCount = 0
    
    foreach ($item in $itemsResponse.value) {
        $fields = $item.fields
        $fileName = $fields.FileLeafRef
        
        if ($fileName -like "*.aspx") {
            $fileCount++
            Write-Log "`nProcessing: $fileName"
            
            $webUrl = "$siteUrl/SitePages/$fileName"
            
            $fileResource = @{
                name        = $fileName
                webUrl      = $webUrl
                description = $fields.Description
            }
            
            $result = ConvertAspxToPdf -fileName $fileName -webUrl $webUrl -fileResource $fileResource
            
            if ($result) {
                $convertedCount++
            }
        }
    }
    
    Write-Log "`n====== CONVERSION SUMMARY ======" -Level SUCCESS
    Write-Log "ASPX files found: $fileCount" -Level INFO
    Write-Log "Files successfully converted: $convertedCount" -Level SUCCESS
    Write-Log "Files failed: $($fileCount - $convertedCount)" -Level $(if ($fileCount -eq $convertedCount) { "SUCCESS" } else { "WARNING" })
    Write-Log "================================" -Level SUCCESS
    
    # Export results to CSV
    if ($global:conversionResults.Count -gt 0) {
        $global:conversionResults | Export-Csv -Path $csvLog -NoTypeInformation -Encoding UTF8
        Write-Log "`nDetailed results exported to: $csvLog" -Level SUCCESS
    }
}

#endregion

#region Main Execution

# Main execution
Write-Log "`n========================================" -Level INFO
Write-Log "ASPX to PDF Conversion - Starting" -Level INFO
Write-Log "========================================" -Level INFO
Write-Log "Site URL: $siteUrl" -Level INFO
Write-Log "Output Folder: $convertedOutputDir" -Level INFO
Write-Log "Log File: $logFile" -Level INFO
Write-Log "CSV Report: $csvLog" -Level INFO
Write-Log "========================================`n" -Level INFO

Get-GraphToken
GetSitePagesDirectly

Write-Log "`n========================================" -Level SUCCESS
Write-Log "Conversion Completed!" -Level SUCCESS
Write-Log "Converted files saved to: $convertedOutputDir" -Level SUCCESS
Write-Log "Log file saved to: $logFile" -Level SUCCESS
if ($global:conversionResults.Count -gt 0) {
    Write-Log "CSV report saved to: $csvLog" -Level SUCCESS
}
Write-Log "========================================" -Level SUCCESS

#endregion
