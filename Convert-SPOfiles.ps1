<#
.SYNOPSIS
    Searches tenant-wide SharePoint and OneDrive locations for Office files and converts them to modern Office formats using Microsoft Graph API.

.DESCRIPTION
    This script authenticates to Microsoft Graph API using client credentials (tenant ID, client ID, and client secret).
    It performs a paginated search query across SharePoint and OneDrive content within the specified region (default is NAM).
    The search targets Office files that can be converted to modern formats including: csv, doc, docx, odp, ods, odt, pot, potm, potx, pps, ppsx, ppsxm, ppt, pptm, pptx, rtf, xls, xlsx.
    Found files are automatically converted to their modern Office format equivalents and saved to a temporary directory.
    Results are exported to a CSV file stored in the user's temporary directory, with a unique timestamped filename.

.PARAMETER tenantId
    The Azure AD tenant ID used for authentication.

.PARAMETER clientId
    The client ID of the Azure AD application used for authentication.

.PARAMETER clientSecret
    The client secret associated with the Azure AD application.

.PARAMETER searchRegion
    The region to scope the search query (default is "NAM").

.PARAMETER searchUrl
    The SharePoint URL/path to search within (default is "https://m365cpi13246019.sharepoint.com/*").

.PARAMETER fileType
    The file types to search for, specified as file extensions (e.g., "docx" or "doc,docx" for multiple types).
    The script automatically adds the "filetype:" prefix for the search query.

.PARAMETER convertTo
    The target format for file conversion. Supported values: "xlsx", "docx", "pptx", "pdf" (default is "docx").
    Note: PDF conversion is supported from modern Office formats. CSV files can only be converted to "xlsx" format.

.PARAMETER AuthType
    The authentication method to use. Supported values: "ClientSecret" or "Certificate" (default is "ClientSecret").

.PARAMETER clientSecret
    The client secret for authentication (used when AuthType = "ClientSecret").

.PARAMETER Thumbprint
    The certificate thumbprint for authentication (used when AuthType = "Certificate").

.PARAMETER CertStore
    The certificate store location. Supported values: "LocalMachine" or "CurrentUser" (default is "LocalMachine").

.PARAMETER downloadFolder
    The local folder path where converted files will be stored (default is the user's temporary directory).

.OUTPUTS
    CSV file containing the following fields for each matching file:
        - ID: The unique identifier of the file.
        - Name: The name of the file.
        - WebURL: The URL to access the file.
        - CreatedDate: The date and time the file was created.
        - LastAccessedDate: The date and time the file was last modified.
        - Owner: The display name of the user who created the file.
        - FileConverted: Whether the file was successfully converted (Yes/No).
        - ConvertedFilePath: The local path where the converted file is saved.
        - SharePointUploadURL: The SharePoint URL where the converted file was uploaded.
    
    Converted files are saved to a timestamped directory in the user's temporary folder and uploaded back to SharePoint.

.NOTES
    Ensure the Azure AD application has appropriate permissions to access Microsoft Graph API and search SharePoint/OneDrive content.
    The script handles pagination automatically, retrieving all available results.
    Files are converted to modern Office formats using Microsoft Graph API's format conversion feature.
    Supported file types for conversion: csv, doc, docx, odp, ods, odt, pot, potm, potx, pps, ppsx, ppsxm, ppt, pptm, pptx, rtf, xls, xlsx.
    Required permissions: Files.Read.All or Files.ReadWrite.All, Sites.Read.All or Sites.ReadWrite.All.

    Authors: Mike Lee
    Modified: 7/18/2025 - Added support for all Office file types and generalized conversion functionality

.LINK
    https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format
    https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview
    https://learn.microsoft.com/en-us/graph/api/driveitem-search?view=graph-rest-1.0&tabs=http
    https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.files/?view=graph-powershell-1.0
    https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.sites/?view=graph-powershell-1.0
    https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users/?view=graph-powershell-1.0
    https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/?view=graph-powershell-1.0

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
    ./Convert-SPOfiles.ps1
    Performs the search for convertible files, converts them to modern Office formats, and exports results to a CSV file in the user's temporary directory.
    Converted files are saved to a timestamped subdirectory.

#>

##############################################################
#                  CONFIGURATION SECTION                    #
#############################################################

# Set the tenant ID, client ID
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3';
$clientId = 'abc64618-283f-47ba-a185-50d935d51d57';

# Authentication configuration
$AuthType = 'Certificate';  # Valid values: 'ClientSecret' or 'Certificate'

# Client Secret authentication (used when $AuthType = 'ClientSecret')
$clientSecret = '';

# Certificate authentication (used when $AuthType = 'Certificate')
$Thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9";  # Certificate thumbprint
$CertStore = 'LocalMachine';  # Valid values: 'LocalMachine' or 'CurrentUser'

# This ensures each log file has a unique name
$date = Get-Date -Format "yyyyMMddHHmmss";

# The log file will store the search results in CSV format
$LogName = Join-Path -Path $env:TEMP -ChildPath ("TenantWide_SharePoint_FileConversion_" + $date + ".csv");

# Search configuration parameters
$searchRegion = "NAM";  # This specifies the region for the search query
$searchUrl = "https://m365cpi13246019.sharepoint.com/sites/SalesandMarketing/Shared%20Documents*";  # This specifies the URL/path to search within
$fileType = "xlsx";  # This specifies the file types to search for (just the extension)
$convertTo = "pdf";  # This specifies the target format for conversion (xlsx, docx, pptx, pdf)

# Download folder configuration
$downloadFolder = $env:TEMP;  # Default to user's temporary directory. Can be set to any valid local path like "C:\Downloads"

#############################################################
#                  CONFIGURATION SECTION                    #
#############################################################

# Validate the convertTo parameter
$supportedFormats = @("xlsx", "docx", "pptx", "pdf")
if ($convertTo -notin $supportedFormats) {
    Write-Host -ForegroundColor Red "Error: Unsupported conversion format '$convertTo'";
    Write-Host -ForegroundColor Yellow "Supported formats are: $($supportedFormats -join ', ')";
    Write-Host -ForegroundColor Yellow "Note: PDF conversion is supported from modern Office formats (docx, xlsx, pptx)";
    exit 1;
}

# Validate the AuthType parameter
$supportedAuthTypes = @("ClientSecret", "Certificate")
if ($AuthType -notin $supportedAuthTypes) {
    Write-Host -ForegroundColor Red "Error: Unsupported authentication type '$AuthType'";
    Write-Host -ForegroundColor Yellow "Supported authentication types are: $($supportedAuthTypes -join ', ')";
    exit 1;
}

# Validate authentication parameters based on AuthType
if ($AuthType -eq 'ClientSecret') {
    if ([string]::IsNullOrWhiteSpace($clientSecret)) {
        Write-Host -ForegroundColor Red "Error: Client secret is required when using ClientSecret authentication";
        exit 1;
    }
}
elseif ($AuthType -eq 'Certificate') {
    if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
        Write-Host -ForegroundColor Red "Error: Certificate thumbprint is required when using Certificate authentication";
        exit 1;
    }
    if ($CertStore -notin @("LocalMachine", "CurrentUser")) {
        Write-Host -ForegroundColor Red "Error: Invalid certificate store '$CertStore'. Valid values are 'LocalMachine' or 'CurrentUser'";
        exit 1;
    }
}

# Validate the downloadFolder parameter
if ([string]::IsNullOrWhiteSpace($downloadFolder)) {
    Write-Host -ForegroundColor Red "Error: Download folder path cannot be empty";
    exit 1;
}

# Create the download folder if it doesn't exist
if (-not (Test-Path $downloadFolder)) {
    try {
        New-Item -ItemType Directory -Path $downloadFolder -Force | Out-Null;
        Write-Host -ForegroundColor Green "Created download folder: $downloadFolder";
    }
    catch {
        Write-Host -ForegroundColor Red "Error: Could not create download folder '$downloadFolder': $($_.Exception.Message)";
        exit 1;
    }
}

# Directory to store converted files
$convertedOutputDir = Join-Path -Path $downloadFolder -ChildPath ("ConvertedFiles_" + $date);
if (-not (Test-Path $convertedOutputDir)) {
    New-Item -ItemType Directory -Path $convertedOutputDir -Force | Out-Null;
}

# Initialize global variables for the token and search results
$global:token = "";
$global:Results = @();

# Check if Excel is available for CSV conversion
Write-Host -ForegroundColor Yellow "Checking for Excel availability for file conversion...";
try {
    $excel = New-Object -ComObject Excel.Application;
    $excel.Quit();
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null;
    Write-Host -ForegroundColor Green "Excel COM object available for file conversion";
    $global:ExcelAvailable = $true;
}
catch {
    Write-Host -ForegroundColor Yellow "Excel COM object not available. Some files may not be convertible.";
    $global:ExcelAvailable = $false;
}

# This function authenticates with Microsoft Graph API and retrieves an access token
function AcquireToken() {
    Write-Host "Connecting to Microsoft Graph using $AuthType authentication..." -ForegroundColor Cyan
    
    if ($AuthType -eq 'ClientSecret') {
        # Client Secret authentication
        $uri = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/v2.0/token";
        
        # Define the body for the authentication request
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = 'https://graph.microsoft.com/.default'
        };
        
        try {
            # Send the authentication request and extract the token
            $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop;
            $global:token = $loginResponse.access_token;
            
            if (-not $global:token) {
                Write-Host -ForegroundColor Red "Authentication failed: No access token received";
                Write-Host -ForegroundColor Yellow "Response: $($loginResponse | ConvertTo-Json -Depth 3)";
                return $false;
            }
            
            Write-Host "Successfully connected using Client Secret authentication" -ForegroundColor Green
            Write-Host -ForegroundColor Cyan "Token length: $($global:token.Length) characters";
            Write-Host -ForegroundColor Cyan "Token starts with: $($global:token.Substring(0, [Math]::Min(20, $global:token.Length)))...";
            return $true;
        }
        catch {
            Write-Host "Failed to connect using Client Secret authentication" -ForegroundColor Red
            Write-Host "Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
            if ($_.Exception.Response) {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream());
                $responseBody = $reader.ReadToEnd();
                Write-Host -ForegroundColor Yellow "Error Response: $responseBody";
            }
            return $false;
        }
    }
    elseif ($AuthType -eq 'Certificate') {
        # Certificate authentication
        $uri = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/v2.0/token";
        
        # Get the certificate from the local certificate store
        try {
            $cert = Get-Item -Path "Cert:\$CertStore\My\$Thumbprint" -ErrorAction Stop
            Write-Host "Certificate found in $CertStore\My store" -ForegroundColor Green
        }
        catch {
            Write-Host "Certificate with thumbprint $Thumbprint not found in $CertStore\My store" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            return $false;
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
            
            if (-not $global:token) {
                Write-Host -ForegroundColor Red "Authentication failed: No access token received";
                Write-Host -ForegroundColor Yellow "Response: $($loginResponse | ConvertTo-Json -Depth 3)";
                return $false;
            }
            
            Write-Host "Successfully connected using Certificate authentication" -ForegroundColor Green
            Write-Host -ForegroundColor Cyan "Token length: $($global:token.Length) characters";
            Write-Host -ForegroundColor Cyan "Token starts with: $($global:token.Substring(0, [Math]::Min(20, $global:token.Length)))...";
            return $true;
        }
        catch {
            Write-Host "Failed to connect using Certificate authentication" -ForegroundColor Red
            Write-Host "Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
            if ($_.Exception.Response) {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream());
                $responseBody = $reader.ReadToEnd();
                Write-Host -ForegroundColor Yellow "Error Response: $responseBody";
            }
            return $false;
        }
    }
    else {
        Write-Host "Invalid authentication type: $AuthType. Valid values are 'ClientSecret' or 'Certificate'." -ForegroundColor Red
        return $false;
    }
}

# This function determines the target format based on the $convertTo parameter
function GetTargetFormat($fileName) {
    $extension = [System.IO.Path]::GetExtension($fileName).ToLower();
    
    # Check if the file is already in the target format
    $targetExtension = ".$convertTo";
    if ($extension -eq $targetExtension) {
        return $null;  # No conversion needed
    }
    
    # Define convertible file extensions
    $convertibleExtensions = @('.csv', '.doc', '.docx', '.odp', '.ods', '.odt', '.pot', '.potm', '.potx', '.pps', '.ppsx', '.ppsxm', '.ppt', '.pptm', '.pptx', '.rtf', '.xls', '.xlsx');
    
    # Check if the file can be converted
    if ($extension -in $convertibleExtensions) {
        return $convertTo;
    }
    
    # File type not supported for conversion
    return $null;
}

# This function converts CSV files to Excel format locally since Graph API doesn't support CSV format conversion
function ConvertCsvToExcelLocally($fileResource, $fileName) {
    try {
        # Define the authorization header
        $headers = @{"Authorization" = "Bearer $global:token" };
        
        # Download the CSV file first
        $csvUrl = $null;
        
        # Get the download URL for the CSV file
        if ($fileResource.parentReference -and $fileResource.parentReference.driveId) {
            $driveId = $fileResource.parentReference.driveId;
            $itemId = $fileResource.id;
            $csvUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/content";
        }
        elseif ($fileResource.id) {
            $itemId = $fileResource.id;
            $csvUrl = "https://graph.microsoft.com/v1.0/drive/items/$itemId/content";
        }
        
        if ($csvUrl) {
            Write-Host -ForegroundColor Gray "Downloading CSV file: $csvUrl";
            
            # Download the CSV content
            $csvContent = Invoke-RestMethod -Uri $csvUrl -Headers $headers -Method GET;
            
            # Save CSV temporarily
            $tempCsvPath = Join-Path -Path $downloadFolder -ChildPath "temp_$fileName";
            $csvContent | Out-File -FilePath $tempCsvPath -Encoding UTF8;
            
            # Convert CSV to the specified target format using Excel COM object
            $targetFormat = GetTargetFormat -fileName $fileName;
            
            # Check if CSV can be converted to the target format
            if ($targetFormat -and $targetFormat -ne "xlsx") {
                Write-Host -ForegroundColor Yellow "CSV files can only be converted to Excel format (xlsx). Target format '$targetFormat' is not supported for CSV files.";
                Write-Host -ForegroundColor Yellow "To convert CSV to PDF, first convert to xlsx, then use a separate tool to convert xlsx to PDF.";
                return $null;
            }
            
            $convertedFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + ".$targetFormat";
            $convertedFilePath = Join-Path -Path $convertedOutputDir -ChildPath $convertedFileName;
            
            try {
                # Only try Excel conversion if Excel is available
                if ($global:ExcelAvailable) {
                    # Create Excel application
                    $excel = New-Object -ComObject Excel.Application;
                    $excel.Visible = $false;
                    $excel.DisplayAlerts = $false;
                    
                    # Open CSV file
                    $workbook = $excel.Workbooks.Open($tempCsvPath);
                    
                    # Save as Excel format (only supported format for CSV)
                    $workbook.SaveAs($convertedFilePath, 51); # 51 = xlOpenXMLWorkbook (.xlsx)
                    
                    # Close and cleanup
                    $workbook.Close();
                    $excel.Quit();
                    
                    # Clean up COM objects
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null;
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null;
                    
                    # Remove temporary CSV file
                    Remove-Item -Path $tempCsvPath -Force;
                    
                    Write-Host -ForegroundColor Green "Successfully converted $fileName to $targetFormat : $convertedFileName";
                    
                    # Upload the converted file back to SharePoint
                    $uploadResult = UploadConvertedFileToSharePoint -convertedFilePath $convertedFilePath -fileResource $fileResource -originalFileName $fileName;
                    
                    # Return both local path and SharePoint URL
                    return @{
                        LocalPath     = $convertedFilePath
                        SharePointURL = $uploadResult
                    }
                }
                else {
                    Write-Host -ForegroundColor Yellow "Excel not available, skipping Excel conversion";
                    throw "Excel not available";
                }
            }
            catch {
                Write-Host -ForegroundColor Red "Excel COM conversion failed: $($_.Exception.Message)";
                Write-Host -ForegroundColor Yellow "Attempting PowerShell-based conversion...";
                
                # Fallback: Simple PowerShell conversion (requires ImportExcel module)
                try {
                    # Check if ImportExcel module is available
                    if (Get-Module -ListAvailable -Name ImportExcel) {
                        Import-Module ImportExcel -ErrorAction Stop;
                        $csvData = Import-Csv -Path $tempCsvPath;
                        $csvData | Export-Excel -Path $convertedFilePath -AutoSize;
                        
                        Remove-Item -Path $tempCsvPath -Force;
                        Write-Host -ForegroundColor Green "Successfully converted $fileName to $targetFormat using ImportExcel: $convertedFileName";
                        
                        # Upload the converted file back to SharePoint
                        $uploadResult = UploadConvertedFileToSharePoint -convertedFilePath $convertedFilePath -fileResource $fileResource -originalFileName $fileName;
                        
                        # Return both local path and SharePoint URL
                        return @{
                            LocalPath     = $convertedFilePath
                            SharePointURL = $uploadResult
                        }
                    }
                    else {
                        Write-Host -ForegroundColor Yellow "ImportExcel module not available. CSV file will be reported but not converted.";
                        # Just copy the CSV file to the output directory for reference
                        $csvCopyPath = Join-Path -Path $convertedOutputDir -ChildPath $fileName;
                        Copy-Item -Path $tempCsvPath -Destination $csvCopyPath;
                        Remove-Item -Path $tempCsvPath -Force;
                        Write-Host -ForegroundColor Yellow "CSV file copied to output directory: $fileName";
                        return $csvCopyPath;
                    }
                }
                catch {
                    Write-Host -ForegroundColor Red "PowerShell conversion also failed: $($_.Exception.Message)";
                    Write-Host -ForegroundColor Yellow "CSV file downloaded but not converted to xlsx";
                    
                    # Clean up temp file
                    if (Test-Path $tempCsvPath) {
                        Remove-Item -Path $tempCsvPath -Force;
                    }
                    return $null;
                }
            }
        }
        else {
            Write-Host -ForegroundColor Red "Could not determine download URL for CSV file: $fileName";
            return $null;
        }
    }
    catch {
        Write-Host -ForegroundColor Red "Error processing CSV file $fileName : $($_.Exception.Message)";
        return $null;
    }
}

# This function uploads a converted file back to the same SharePoint library
function UploadConvertedFileToSharePoint($convertedFilePath, $fileResource, $originalFileName) {
    try {
        if (-not (Test-Path $convertedFilePath)) {
            Write-Host -ForegroundColor Red "Converted file not found: $convertedFilePath";
            return $false;
        }
        
        # Define the authorization header
        $headers = @{"Authorization" = "Bearer $global:token" };
        
        # Get the parent folder information
        $driveId = $fileResource.parentReference.driveId;
        $parentFolderId = $fileResource.parentReference.id;
        
        if (-not $driveId -or -not $parentFolderId) {
            Write-Host -ForegroundColor Red "Could not determine parent folder for upload";
            return $false;
        }
        
        # Get the converted filename
        $convertedFileName = [System.IO.Path]::GetFileName($convertedFilePath);
        
        # Read the file content
        $fileContent = [System.IO.File]::ReadAllBytes($convertedFilePath);
        
        # Create the upload URL
        $uploadUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$parentFolderId`:/$convertedFileName`:/content";
        
        # Determine content type based on file extension
        $extension = [System.IO.Path]::GetExtension($convertedFileName).ToLower();
        $contentType = switch ($extension) {
            '.xlsx' { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
            '.docx' { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
            '.pptx' { 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }
            '.pdf' { 'application/pdf' }
            default { 'application/octet-stream' }
        }
        
        Write-Host -ForegroundColor Cyan "Uploading $convertedFileName to SharePoint...";
        Write-Host -ForegroundColor Gray "Upload URL: $uploadUrl";
        
        # Upload the file
        $response = Invoke-RestMethod -Uri $uploadUrl -Method PUT -Headers $headers -Body $fileContent -ContentType $contentType;
        
        if ($response.id) {
            Write-Host -ForegroundColor Green "Successfully uploaded $convertedFileName to SharePoint";
            Write-Host -ForegroundColor Gray "New file ID: $($response.id)";
            Write-Host -ForegroundColor Gray "SharePoint URL: $($response.webUrl)";
            return $response.webUrl;
        }
        else {
            Write-Host -ForegroundColor Red "Upload response did not contain file ID";
            return $false;
        }
    }
    catch {
        Write-Host -ForegroundColor Red "Error uploading $convertedFileName to SharePoint: $($_.Exception.Message)";
        if ($_.Exception.Response) {
            Write-Host -ForegroundColor Red "HTTP Status: $($_.Exception.Response.StatusCode)";
            Write-Host -ForegroundColor Red "HTTP Status Description: $($_.Exception.Response.StatusDescription)";
            
            # Try to read the response body for more details
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream());
                $responseBody = $reader.ReadToEnd();
                Write-Host -ForegroundColor Yellow "Upload Error Response: $responseBody";
            }
            catch {
                Write-Host -ForegroundColor Yellow "Could not read error response body";
            }
        }
        return $false;
    }
}

# This function converts a file to modern Office format using Microsoft Graph API
function ConvertToModernFormat($fileResource, $fileName) {
    try {
        # Define the authorization header
        $headers = @{"Authorization" = "Bearer $global:token" };
        
        # Get the file extension and determine target format
        $extension = [System.IO.Path]::GetExtension($fileName).ToLower();
        $targetFormat = GetTargetFormat -fileName $fileName;
        
        # Check if conversion is needed
        if (-not $targetFormat) {
            Write-Host -ForegroundColor Yellow "File $fileName is already in modern format or not supported for conversion";
            return $null;
        }
        
        # Handle CSV files separately (Graph API doesn't support CSV conversion)
        if ($extension -eq '.csv') {
            Write-Host -ForegroundColor Yellow "CSV files cannot be converted via Graph API - attempting local conversion...";
            return ConvertCsvToExcelLocally -fileResource $fileResource -fileName $fileName;
        }
        
        # Define supported extensions for Graph API conversion
        $graphApiSupportedExtensions = @('.doc', '.docx', '.odp', '.ods', '.odt', '.pot', '.potm', '.potx', '.pps', '.ppsx', '.ppsxm', '.ppt', '.pptm', '.pptx', '.rtf', '.xls', '.xlsx');
        
        if ($extension -in $graphApiSupportedExtensions) {
            Write-Host -ForegroundColor Cyan "Converting $fileName to $targetFormat...";
            
            # Try different approaches to get the correct endpoint
            $convertUrl = $null;
            
            # First, try using the parentReference and name
            if ($fileResource.parentReference -and $fileResource.parentReference.driveId) {
                $driveId = $fileResource.parentReference.driveId;
                $itemId = $fileResource.id;
                $convertUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/content?format=$targetFormat";
            }
            # Fallback to direct ID approach
            elseif ($fileResource.id) {
                $itemId = $fileResource.id;
                $convertUrl = "https://graph.microsoft.com/v1.0/drive/items/$itemId/content?format=$targetFormat";
            }
            
            if ($convertUrl) {
                Write-Host -ForegroundColor Gray "Using conversion URL: $convertUrl";
                
                # Make the conversion request - handle 302 redirect properly
                try {
                    $response = Invoke-WebRequest -Uri $convertUrl -Headers $headers -Method GET -MaximumRedirection 0;
                }
                catch {
                    # Check if it's a 302 redirect (which is expected)
                    if ($_.Exception.Response.StatusCode -eq 302) {
                        $response = $_.Exception.Response;
                    }
                    else {
                        throw;
                    }
                }
                
                if ($response.StatusCode -eq 302) {
                    # Get the redirect URL for the converted file
                    $downloadUrl = $response.Headers.Location;
                    
                    if (-not $downloadUrl) {
                        # Try different ways to get the location header
                        $downloadUrl = $response.Headers["Location"];
                        if (-not $downloadUrl -and $response.Headers.GetEnumerator()) {
                            foreach ($header in $response.Headers.GetEnumerator()) {
                                if ($header.Key -eq "Location") {
                                    $downloadUrl = $header.Value;
                                    break;
                                }
                            }
                        }
                    }
                    
                    if ($downloadUrl) {
                        Write-Host -ForegroundColor Gray "Download URL: $downloadUrl";
                        
                        # Download the converted file
                        $convertedFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + ".$targetFormat";
                        $convertedFilePath = Join-Path -Path $convertedOutputDir -ChildPath $convertedFileName;
                        
                        Invoke-WebRequest -Uri $downloadUrl -OutFile $convertedFilePath;
                        Write-Host -ForegroundColor Green "Successfully converted $fileName to $targetFormat : $convertedFileName";
                        
                        # Upload the converted file back to SharePoint
                        $uploadResult = UploadConvertedFileToSharePoint -convertedFilePath $convertedFilePath -fileResource $fileResource -originalFileName $fileName;
                        
                        # Return both local path and SharePoint URL
                        return @{
                            LocalPath     = $convertedFilePath
                            SharePointURL = $uploadResult
                        }
                    }
                    else {
                        Write-Host -ForegroundColor Red "No download URL found in 302 response";
                        return $null;
                    }
                }
                else {
                    Write-Host -ForegroundColor Red "Conversion failed with status code: $($response.StatusCode)";
                    return $null;
                }
            }
            else {
                Write-Host -ForegroundColor Red "Could not determine conversion URL for $fileName";
                return $null;
            }
        }
        else {
            Write-Host -ForegroundColor Yellow "File $fileName with extension $extension is not supported for conversion";
            return $null;
        }
    }
    catch {
        Write-Host -ForegroundColor Red "Error converting $fileName`: $($_.Exception.Message)";
        # Add more detailed error information
        if ($_.Exception.Response) {
            Write-Host -ForegroundColor Red "HTTP Status: $($_.Exception.Response.StatusCode)";
            Write-Host -ForegroundColor Red "HTTP Status Description: $($_.Exception.Response.StatusDescription)";
        }
        return $null;
    }
}

# This function formats the file type parameter into a proper search query format
function Format-FileTypeQuery($fileTypeInput) {
    if ([string]::IsNullOrWhiteSpace($fileTypeInput)) {
        return "";
    }
    
    Write-Host -ForegroundColor Gray "Input file type: '$fileTypeInput'";
    
    # Split by comma and trim whitespace
    $extensions = $fileTypeInput.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    
    Write-Host -ForegroundColor Gray "Extensions found: $($extensions -join ', ')";
    
    # Format each extension with "filetype:" prefix
    $formattedExtensions = @()
    foreach ($ext in $extensions) {
        if ($ext.StartsWith("filetype:")) {
            $formattedExtensions += $ext  # Already formatted
        }
        else {
            $formattedExtensions += "filetype:$ext"
        }
    }
    
    Write-Host -ForegroundColor Gray "Formatted extensions: $($formattedExtensions -join ', ')";
    
    # Join with OR operator if multiple extensions
    if ($formattedExtensions.Count -gt 1) {
        $result = "(" + ($formattedExtensions -join " OR ") + ")"
    }
    else {
        $result = $formattedExtensions[0]
    }
    
    Write-Host -ForegroundColor Gray "Final formatted query: '$result'";
    return $result;
}

# This function sends a search request to Microsoft Graph API and handles pagination
function PerformSearch {
    # Display the search query information
    Write-Host -ForegroundColor Green "Performing Search for Office Files and Converting to Modern Formats";
    
    # Define the authorization header
    $headers = @{"Authorization" = "Bearer $global:token" };
    $string = "https://graph.microsoft.com/v1.0/search/query"; 
    
    Write-Host -ForegroundColor Cyan "Token for search request: $($global:token.Substring(0, [Math]::Min(20, $global:token.Length)))...";
    Write-Host -ForegroundColor Cyan "Authorization header: Bearer $($global:token.Substring(0, [Math]::Min(20, $global:token.Length)))..."; 

    # Initialize variables for pagination
    $moreresults = $true;
    $start = 0;
    $size = 200;
    $i = 0;

    # Loop to handle pagination
    while ($moreresults) {
        # Format the file type for the search query
        $formattedFileType = Format-FileTypeQuery -fileTypeInput $fileType
        
        # The query searches for Office files that can be converted to modern formats in the specified region
        $requestPayload = @"
        {
            "requests": [
                {
                    "entityTypes": [  
                    "driveItem"
                    ],
                    "query": {
                        "queryString": "$formattedFileType (path:\"$searchUrl\")"
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
        try {
            Write-Host -ForegroundColor Gray "Sending search request batch $($i + 1)...";
            $Results = Invoke-RestMethod -Method POST -Uri $string -Headers $headers -Body $requestPayload -ContentType "application/json";
            Write-Host -ForegroundColor Gray "Search request successful";
        }
        catch {
            Write-Host -ForegroundColor Red "Search request failed: $($_.Exception.Message)";
            if ($_.Exception.Response) {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream());
                $responseBody = $reader.ReadToEnd();
                Write-Host -ForegroundColor Yellow "Search Error Response: $responseBody";
            }
            break;
        }

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

    Write-Host -ForegroundColor Green "Search and File Conversion Completed Successfully";
    Write-Host ""
    Write-Host -ForegroundColor Yellow "Results Exported to $logName";
    Write-Host -ForegroundColor Yellow "Converted files saved to $convertedOutputDir";
}


# This function extracts relevant fields from the search results, converts files to modern formats, and appends them to the CSV file
function ExportResultSet($results) {
    $Results.value.hitsContainers.hits.resource | ForEach-Object {
        $fileResource = $_;
        $fileId = $_.id;
        $fileName = $_.name;
        $webUrl = $_.webUrl;
        $createdDate = $_.createdDateTime;
        $lastModifiedDate = $_.lastModifiedDateTime;
        $owner = $_.createdBy.user.displayName;
        
        # Convert the file to modern format
        $conversionResult = ConvertToModernFormat -fileResource $fileResource -fileName $fileName;
        
        # Handle the result format (could be a hashtable with LocalPath and SharePointURL, or just a string path)
        $convertedFilePath = $null;
        $sharePointURL = "N/A";
        
        if ($conversionResult) {
            if ($conversionResult -is [hashtable]) {
                $convertedFilePath = $conversionResult.LocalPath;
                $sharePointURL = if ($conversionResult.SharePointURL) { $conversionResult.SharePointURL } else { "Upload failed" };
            }
            else {
                $convertedFilePath = $conversionResult;
                $sharePointURL = "N/A";
            }
        }
        
        # Create the result object
        $resultObject = [PSCustomObject]@{
            ID                  = $fileId
            Name                = $fileName
            WebURL              = $webUrl
            CreatedDate         = $createdDate
            LastAccessedDate    = $lastModifiedDate
            Owner               = $owner
            FileConverted       = if ($convertedFilePath) { "Yes" } else { "No" }
            ConvertedFilePath   = if ($convertedFilePath) { $convertedFilePath } else { "N/A" }
            SharePointUploadURL = $sharePointURL
        }
        
        # Export to CSV
        $resultObject | Export-Csv $logName -NoTypeInformation -NoClobber -Append;
    }
}

# This is the first step before performing any search queries
$authResult = AcquireToken;

# Only proceed if authentication was successful
if ($authResult) {
    # Test the token with a simple Graph API call
    Write-Host -ForegroundColor Yellow "Testing token with a simple Graph API call...";
    try {
        $testHeaders = @{"Authorization" = "Bearer $global:token" };
        $null = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?search=*" -Headers $testHeaders -Method GET;
        Write-Host -ForegroundColor Green "Token test successful - proceeding with search";
    }
    catch {
        Write-Host -ForegroundColor Red "Token test failed: $($_.Exception.Message)";
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream());
            $responseBody = $reader.ReadToEnd();
            Write-Host -ForegroundColor Yellow "Test Error Response: $responseBody";
        }
        Write-Host -ForegroundColor Red "Script terminated due to token validation failure";
        exit 1;
    }
    
    # Perform search for each query
    PerformSearch;
}
else {
    Write-Host -ForegroundColor Red "Script terminated due to authentication failure";
    exit 1;
} 
