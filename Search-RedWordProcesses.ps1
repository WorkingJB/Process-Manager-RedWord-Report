<#
.SYNOPSIS
    Searches Process Manager for processes containing "red flag" words and generates a detailed CSV report.

.DESCRIPTION
    This script authenticates to Process Manager, searches for processes containing specified red flag words,
    retrieves detailed information about each process, and exports the results to a CSV file.

.EXAMPLE
    .\Search-RedWordProcesses.ps1

.NOTES
    Author: Process Manager Red Word Search Tool
    Version: 1.2

.CHANGELOG
    v1.2 - Enhanced authentication with detailed debugging output
         - Added support for any Process Manager URL (not just demo)
         - Improved error handling with clear user messages
         - Added "Press any key to exit" to prevent PowerShell auto-close
         - Better Search API authentication with fallback to main token
         - Added verbose logging for all API calls
         - Improved User ID and Tenant ID extraction with error handling
    v1.1 - Fixed URL parsing to support both base URLs and full tenant URLs
         - Fixed variable name conflict with PowerShell $host variable
         - Improved tenant ID extraction logic
    v1.0 - Initial release
#>

#Requires -Version 5.1

# Regional endpoint mapping
$RegionalEndpoints = @{
    'demo.promapp.com' = 'https://dmo-wus-sch.promapp.io'
    'us.promapp.com' = 'https://prd-wus-sch.promapp.io'
    'ca.promapp.com' = 'https://prd-cac-sch.promapp.io'
    'eu.promapp.com' = 'https://prd-neu-sch.promapp.io'
    'au.promapp.com' = 'https://prd-aus-sch.promapp.io'
}

# Function to parse Process Manager URL and extract base URL and tenant ID
function Parse-ProcessManagerUrl {
    param([string]$InputUrl)

    # Validate and clean URL
    if ($InputUrl -notmatch '^https?://') {
        $InputUrl = "https://$InputUrl"
    }
    $InputUrl = $InputUrl.TrimEnd('/')

    $uri = [System.Uri]$InputUrl
    $baseUrl = "$($uri.Scheme)://$($uri.Host)"

    # Try to extract tenant ID from URL path
    $tenantId = $null
    if ($uri.AbsolutePath -match '^/([a-f0-9]{32})') {
        $tenantId = $matches[1]
        Write-Verbose "Extracted tenant ID from URL: $tenantId"
    }

    return @{
        BaseUrl = $baseUrl
        TenantId = $tenantId
        FullUrl = $InputUrl
        Hostname = $uri.Host
    }
}

# Function to get credentials
function Get-ProcessManagerCredentials {
    Write-Host "`n=== Process Manager Red Word Search Tool ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "You can enter either:"
    Write-Host "  - Base URL: https://demo.promapp.com"
    Write-Host "  - Full URL with tenant: https://demo.promapp.com/93555a16ceb24f139a6e8a40618d3f8b"
    Write-Host ""

    # Get URL
    $url = Read-Host "Enter your Process Manager URL"

    # Parse the URL
    $urlInfo = Parse-ProcessManagerUrl -InputUrl $url

    # Get credentials
    Write-Host ""
    $username = Read-Host "Enter your username"
    $securePassword = Read-Host "Enter your password" -AsSecureString
    $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))

    return @{
        BaseUrl = $urlInfo.BaseUrl
        TenantId = $urlInfo.TenantId
        Hostname = $urlInfo.Hostname
        Username = $username
        Password = $password
    }
}

# Function to determine search endpoint from hostname
function Get-SearchEndpoint {
    param([string]$Hostname)

    if ($RegionalEndpoints.ContainsKey($Hostname)) {
        return $RegionalEndpoints[$Hostname]
    }

    Write-Warning "Unknown region for hostname: $Hostname. Using Demo region endpoint."
    return $RegionalEndpoints['demo.promapp.com']
}

# Function to extract tenant ID from URL
function Get-TenantId {
    param([string]$BaseUrl, [string]$AccessToken)

    # The tenant ID is typically in the JWT token
    try {
        $tokenParts = $AccessToken.Split('.')
        if ($tokenParts.Count -ge 2) {
            $payload = $tokenParts[1]
            # Add padding if needed
            $padding = (4 - ($payload.Length % 4)) % 4
            $payload += "=" * $padding

            $decodedBytes = [System.Convert]::FromBase64String($payload)
            $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
            $tokenData = $decodedJson | ConvertFrom-Json

            if ($tokenData.TenantName) {
                return $tokenData.TenantName
            }
        }
    }
    catch {
        Write-Verbose "Could not extract tenant from token: $_"
    }

    return $null
}

# Function to authenticate with Process Manager
function Get-ProcessManagerToken {
    param(
        [string]$BaseUrl,
        [string]$Username,
        [string]$Password,
        [string]$TenantId = $null
    )

    Write-Host "`nAuthenticating to Process Manager..." -ForegroundColor Yellow

    $body = @{
        grant_type = 'password'
        username = $Username
        password = $Password
    }

    # If tenant ID is provided, try that first
    if ($TenantId) {
        $authUrl = "$BaseUrl/$TenantId/oauth2/token"
        Write-Host "  Trying: $authUrl" -ForegroundColor Gray

        try {
            $response = Invoke-RestMethod -Uri $authUrl -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded'

            if ($response.access_token) {
                Write-Host "  Authentication successful!" -ForegroundColor Green
                return @{
                    AccessToken = $response.access_token
                    TokenType = $response.token_type
                    ExpiresIn = $response.expires_in
                    TenantId = $TenantId
                }
            }
        }
        catch {
            Write-Host "  Failed with tenant ID from URL: $($_.Exception.Message)" -ForegroundColor Gray
        }
    }

    # Try without tenant ID (some instances support this)
    $authUrl = "$BaseUrl/oauth2/token"
    Write-Host "  Trying: $authUrl" -ForegroundColor Gray

    try {
        $response = Invoke-RestMethod -Uri $authUrl -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded'

        if ($response.access_token) {
            Write-Host "  Authentication successful!" -ForegroundColor Green
            return @{
                AccessToken = $response.access_token
                TokenType = $response.token_type
                ExpiresIn = $response.expires_in
            }
        }
    }
    catch {
        Write-Host "  Failed without tenant ID: $($_.Exception.Message)" -ForegroundColor Gray
    }

    # Try to extract tenant from the main page
    Write-Host "  Attempting to discover tenant ID from main page..." -ForegroundColor Gray
    try {
        $mainPage = Invoke-WebRequest -Uri $BaseUrl -UseBasicParsing
        if ($mainPage.Content -match '/([a-f0-9]{32})/') {
            $discoveredTenantId = $matches[1]
            Write-Host "  Found tenant ID: $discoveredTenantId" -ForegroundColor Cyan

            $authUrl = "$BaseUrl/$discoveredTenantId/oauth2/token"
            Write-Host "  Trying: $authUrl" -ForegroundColor Gray

            $response = Invoke-RestMethod -Uri $authUrl -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded'

            if ($response.access_token) {
                Write-Host "  Authentication successful!" -ForegroundColor Green
                return @{
                    AccessToken = $response.access_token
                    TokenType = $response.token_type
                    ExpiresIn = $response.expires_in
                    TenantId = $discoveredTenantId
                }
            }
        }
        else {
            Write-Host "  Could not find tenant ID in page content" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "  Failed to discover tenant ID: $($_.Exception.Message)" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "ERROR: Could not authenticate to Process Manager" -ForegroundColor Red
    Write-Host "Please verify:" -ForegroundColor Yellow
    Write-Host "  1. Your URL is correct (e.g., https://demo.promapp.com or https://demo.promapp.com/tenant-id)" -ForegroundColor Yellow
    Write-Host "  2. Your username and password are correct" -ForegroundColor Yellow
    Write-Host "  3. You have network access to the Process Manager instance" -ForegroundColor Yellow
    Write-Host ""
    return $null
}

# Function to authenticate with Search API
function Get-SearchToken {
    param(
        [string]$SearchEndpoint,
        [string]$TenantId,
        [int]$UserId,
        [string]$AccessToken
    )

    Write-Host "`nAuthenticating to Search API..." -ForegroundColor Yellow

    $authUrl = "$SearchEndpoint/api/authentication/$TenantId/$UserId"
    Write-Host "  Trying: $authUrl" -ForegroundColor Gray
    Write-Host "  Tenant ID: $TenantId" -ForegroundColor Gray
    Write-Host "  User ID: $UserId" -ForegroundColor Gray

    $headers = @{
        'Authorization' = "Bearer $AccessToken"
        'Content-Type' = 'application/json'
    }

    try {
        $response = Invoke-RestMethod -Uri $authUrl -Method Post -Headers $headers

        if ($response.Status -eq 'Success' -and $response.Message) {
            Write-Host "  Search API authentication successful!" -ForegroundColor Green
            return $response.Message
        }
        else {
            Write-Host "  Unexpected response from Search API" -ForegroundColor Yellow
            Write-Host "  Response: $($response | ConvertTo-Json -Depth 2)" -ForegroundColor Gray
            Write-Host "  Will attempt to use main access token for search..." -ForegroundColor Yellow
            return $AccessToken
        }
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $statusDescription = $_.Exception.Response.StatusDescription

        Write-Host "  Search API authentication failed" -ForegroundColor Yellow
        Write-Host "  Status Code: $statusCode - $statusDescription" -ForegroundColor Gray
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Gray

        if ($statusCode -eq 404) {
            Write-Host ""
            Write-Host "  NOTE: 404 error usually means:" -ForegroundColor Yellow
            Write-Host "    - The tenant ID or user ID is incorrect" -ForegroundColor Yellow
            Write-Host "    - The search endpoint URL is wrong for your region" -ForegroundColor Yellow
        }

        Write-Host ""
        Write-Host "  Will attempt to use main access token for search..." -ForegroundColor Cyan
        return $AccessToken
    }
}

# Function to search for processes
function Search-Processes {
    param(
        [string]$SearchEndpoint,
        [string]$SearchToken,
        [string]$SearchTerm,
        [int]$PageNumber = 1
    )

    $encodedTerm = [System.Web.HttpUtility]::UrlEncode("`"$SearchTerm`"")
    $searchUrl = "$SearchEndpoint/fullsearch?SearchCriteria=$encodedTerm&IncludedTypes=1&SearchMatchType=0&pageNumber=$PageNumber"

    $headers = @{
        'Authorization' = "Bearer $SearchToken"
        'Content-Type' = 'application/json'
    }

    try {
        $response = Invoke-RestMethod -Uri $searchUrl -Method Get -Headers $headers
        return $response
    }
    catch {
        Write-Warning "Search failed for term '$SearchTerm': $($_.Exception.Message)"
        return $null
    }
}

# Function to get process details
function Get-ProcessDetails {
    param(
        [string]$BaseUrl,
        [string]$TenantId,
        [string]$ProcessId,
        [string]$AccessToken
    )

    $apiUrl = "$BaseUrl/$TenantId/api/Process/$ProcessId"

    $headers = @{
        'Authorization' = "Bearer $AccessToken"
        'Content-Type' = 'application/json'
    }

    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
        return $response
    }
    catch {
        Write-Warning "Failed to get process details for ID $ProcessId : $($_.Exception.Message)"
        return $null
    }
}

# Function to determine process status
function Get-ProcessStatus {
    param(
        [string]$EntityType,
        [string]$State
    )

    if ($EntityType -eq 'PublishedProcess') {
        return 'Published'
    }
    elseif ($EntityType -eq 'UnpublishedProcess') {
        if ($State -eq 'Draft') {
            return 'In Progress'
        }
        return 'Unpublished'
    }

    return 'Unknown'
}

# Function to get red words from user
function Get-RedWords {
    Write-Host "`n=== Red Flag Words Configuration ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "You can either:"
    Write-Host "  1. Enter red flag words manually (comma-separated)"
    Write-Host "  2. Load from a text file (one word per line)"
    Write-Host ""

    $choice = Read-Host "Enter your choice (1 or 2)"

    if ($choice -eq '2') {
        $filePath = Read-Host "Enter the path to the text file"

        if (Test-Path $filePath) {
            $words = Get-Content $filePath | Where-Object {
                $line = $_.Trim()
                $line -ne '' -and -not $line.StartsWith('#')
            }
            Write-Host "Loaded $($words.Count) red flag words from file." -ForegroundColor Green
            return $words
        }
        else {
            Write-Warning "File not found. Please enter words manually."
        }
    }

    # Manual entry
    Write-Host ""
    Write-Host "Enter red flag words (comma-separated):"
    $input = Read-Host
    $words = $input -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

    Write-Host "Loaded $($words.Count) red flag words." -ForegroundColor Green
    return $words
}

# Main execution
function Main {
    # Add required assembly for URL encoding
    Add-Type -AssemblyName System.Web

    # Get credentials
    $credentials = Get-ProcessManagerCredentials

    # Determine search endpoint
    $searchEndpoint = Get-SearchEndpoint -Hostname $credentials.Hostname
    Write-Host "`nUsing search endpoint: $searchEndpoint" -ForegroundColor Cyan

    # Authenticate to Process Manager
    $authResult = Get-ProcessManagerToken -BaseUrl $credentials.BaseUrl -Username $credentials.Username -Password $credentials.Password -TenantId $credentials.TenantId

    if (-not $authResult -or -not $authResult.AccessToken) {
        Write-Host ""
        Write-Host "ERROR: Authentication failed. Cannot continue." -ForegroundColor Red
        Write-Host ""
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }

    # Determine tenant ID (from URL if provided, otherwise from token or auth result)
    $tenantId = $credentials.TenantId

    if (-not $tenantId) {
        # Try to extract from token
        $tenantId = Get-TenantId -BaseUrl $credentials.BaseUrl -AccessToken $authResult.AccessToken
    }

    if (-not $tenantId) {
        # Try to extract from auth result if available
        if ($authResult.TenantId) {
            $tenantId = $authResult.TenantId
        }
        else {
            Write-Host ""
            Write-Host "ERROR: Could not determine tenant ID." -ForegroundColor Red
            Write-Host "Please provide the full URL including the tenant ID." -ForegroundColor Yellow
            Write-Host "Example: https://demo.promapp.com/93555a16ceb24f139a6e8a40618d3f8b" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Press any key to exit..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
    }

    Write-Host ""
    Write-Host "Tenant ID: $tenantId" -ForegroundColor Cyan

    # Get user ID from token
    $userId = $null
    try {
        $tokenParts = $authResult.AccessToken.Split('.')
        if ($tokenParts.Count -ge 2) {
            $payload = $tokenParts[1]
            $padding = (4 - ($payload.Length % 4)) % 4
            $payload += "=" * $padding
            $decodedBytes = [System.Convert]::FromBase64String($payload)
            $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
            $tokenData = $decodedJson | ConvertFrom-Json
            $userId = $tokenData.UserId
            Write-Host "User ID: $userId" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Warning "Could not extract User ID from token: $($_.Exception.Message)"
    }

    if (-not $userId) {
        Write-Host ""
        Write-Host "WARNING: Could not determine User ID from token." -ForegroundColor Yellow
        Write-Host "This may affect search API authentication." -ForegroundColor Yellow
        Write-Host ""
    }

    # Authenticate to Search API
    $searchToken = Get-SearchToken -SearchEndpoint $searchEndpoint -TenantId $tenantId -UserId $userId -AccessToken $authResult.AccessToken

    if (-not $searchToken) {
        Write-Host ""
        Write-Host "ERROR: Search API authentication failed and could not fall back to main token." -ForegroundColor Red
        Write-Host ""
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }

    # Get red flag words
    $redWords = Get-RedWords

    if ($redWords.Count -eq 0) {
        Write-Host ""
        Write-Host "ERROR: No red flag words specified." -ForegroundColor Red
        Write-Host ""
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }

    # Search for processes
    Write-Host "`n=== Searching for Processes ===" -ForegroundColor Cyan
    $allResults = @()
    $processCache = @{}

    foreach ($word in $redWords) {
        Write-Host "`nSearching for: '$word'" -ForegroundColor Yellow

        $pageNumber = 1
        $hasMorePages = $true

        while ($hasMorePages) {
            $searchResult = Search-Processes -SearchEndpoint $searchEndpoint -SearchToken $searchToken -SearchTerm $word -PageNumber $pageNumber

            if ($searchResult -and $searchResult.success -and $searchResult.response) {
                Write-Host "  Found $($searchResult.response.Count) processes on page $pageNumber" -ForegroundColor Gray

                foreach ($process in $searchResult.response) {
                    $processId = $process.ProcessUniqueId

                    # Check if we already processed this process
                    if (-not $processCache.ContainsKey($processId)) {
                        # Get detailed process information
                        Write-Verbose "Getting details for process: $($process.Name)"
                        $processDetails = Get-ProcessDetails -BaseUrl $credentials.BaseUrl -TenantId $tenantId -ProcessId $processId -AccessToken $authResult.AccessToken

                        # Cache the process
                        $processCache[$processId] = @{
                            SearchResult = $process
                            Details = $processDetails
                            RedWords = @($word)
                        }
                    }
                    else {
                        # Add this red word to the existing entry
                        if ($processCache[$processId].RedWords -notcontains $word) {
                            $processCache[$processId].RedWords += $word
                        }
                    }
                }

                # Check if there are more pages
                if ($searchResult.paging -and -not $searchResult.paging.IsLastPage) {
                    $pageNumber++
                }
                else {
                    $hasMorePages = $false
                }
            }
            else {
                $hasMorePages = $false
            }
        }
    }

    Write-Host "`n=== Compiling Results ===" -ForegroundColor Cyan
    Write-Host "Found $($processCache.Count) unique processes" -ForegroundColor Green

    # Compile results
    foreach ($entry in $processCache.GetEnumerator()) {
        $searchData = $entry.Value.SearchResult
        $detailsData = $entry.Value.Details
        $redWordsFound = $entry.Value.RedWords -join ', '

        # Extract variation name if present
        $variationName = ''
        if ($detailsData -and $detailsData.variationSetData) {
            $variationName = $detailsData.variationSetData.VariationName
        }

        # Get owner and expert
        $owner = ''
        $expert = ''
        if ($detailsData -and $detailsData.processJson) {
            $owner = $detailsData.processJson.Owner
            $expert = $detailsData.processJson.Expert
        }

        # Get process group path
        $groupPath = ''
        if ($searchData.BreadCrumbGroupNames) {
            $groupPath = $searchData.BreadCrumbGroupNames -join ' > '
        }

        # Get status
        $status = Get-ProcessStatus -EntityType $searchData.EntityType -State $(if ($detailsData -and $detailsData.processJson) { $detailsData.processJson.State } else { '' })

        # Get publish date (we need to check if there's publish info in the details)
        $publishDate = ''
        if ($detailsData -and $detailsData.processJson -and $detailsData.processJson.PublishDate) {
            $publishDate = $detailsData.processJson.PublishDate
        }

        # Get review status (this might not be in the API response, marking as N/A for now)
        $reviewStatus = 'N/A'
        if ($detailsData -and $detailsData.processJson -and $detailsData.processJson.ReviewStatus) {
            $reviewStatus = $detailsData.processJson.ReviewStatus
        }

        # Create result object
        $result = [PSCustomObject]@{
            'Process Title' = $searchData.Name
            'Process Variation Name' = $variationName
            'Red Flag Words' = $redWordsFound
            'Process Owner' = $owner
            'Process Expert' = $expert
            'Process Group Path' = $groupPath
            'Status' = $status
            'Publish Date' = $publishDate
            'Review Status' = $reviewStatus
            'Process URL' = $searchData.ItemUrl
            'Process ID' = $searchData.ProcessUniqueId
        }

        $allResults += $result
    }

    # Export to CSV
    if ($allResults.Count -gt 0) {
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $outputFile = "RedWordProcesses_$timestamp.csv"

        $allResults | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

        Write-Host "`n=== Export Complete ===" -ForegroundColor Green
        Write-Host "Results exported to: $outputFile" -ForegroundColor Cyan
        Write-Host "Total processes found: $($allResults.Count)" -ForegroundColor Cyan

        # Display summary
        Write-Host "`n=== Summary ===" -ForegroundColor Cyan
        Write-Host "Processes by status:"
        $allResults | Group-Object Status | ForEach-Object {
            Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor Gray
        }
    }
    else {
        Write-Host "`nNo processes found containing the specified red flag words." -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Run the script with error handling
try {
    Main
}
catch {
    Write-Host ""
    Write-Host "FATAL ERROR: An unexpected error occurred" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Stack Trace:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
