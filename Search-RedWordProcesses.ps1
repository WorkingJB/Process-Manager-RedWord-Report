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
    Version: 1.8

.CHANGELOG
    v1.8 - Added enhanced data extraction features
         - Parse variation names from process titles (splits on "::" delimiter)
         - Extract published date from Published.PublishedDate object (ISO format)
         - Added review due date API call to /{tenant}/Api/v1/Processes/{id}/ReviewDue
         - Calculate review status (In Date/Out of Date) based on review due date
         - Added Review Due Date column to CSV output
         - Format all dates as yyyy-MM-dd for consistency
         - Process titles now separated from variation names in output
    v1.7 - CRITICAL FIX: Corrected process details API endpoint
         - Fixed URL from /api/Process/ to /Api/v1/Processes/
         - Changes: capital 'A' in Api, added v1, made Processes plural
         - Verified against API spec: /{tenant}/Api/v1/Processes/{processId}
         - Confirmed using correct token (site auth token, not search token)
         - Enhanced error output showing URL and status code for debugging
         - Removed unnecessary Content-Type header from GET request
    v1.6 - CRITICAL FIX: Removed ProcessSearchFields parameters causing 404 errors
         - Reviewed working implementation - ProcessSearchFields not required
         - Simplified search URL to match reference implementation
         - Removed ProcessSearchFields=1,2,3,4 from query string
         - Now uses API default field matching behavior
         - Added detailed debugging: shows exact URL and token being used
         - Removed unnecessary Content-Type header from search GET request
    v1.5 - CRITICAL FIX: Changed search token request from POST to GET
         - Reviewed reference implementation in UnpublishedProcessDocuments repo
         - Search token endpoint requires GET request, not POST
         - Removed unnecessary Content-Type header from GET request
         - Added access token preview in debug output
    v1.4 - CRITICAL FIX: Corrected search token authentication endpoint
         - Search token now requested from main site: {BaseUrl}/{TenantId}/search/GetSearchServiceToken
         - Removed incorrect authentication to regional search endpoint
         - Removed User ID requirement (not needed for search service token)
         - Searches now use correct flow: main site auth → get search token → use on regional endpoint
    v1.3 - Fixed Search API query parameters to include all ProcessSearchFields (1,2,3,4)
         - Ensured search terms are properly quoted for literal matching vs fuzzy
         - Added token type indicator (dedicated search token vs fallback)
         - Enhanced search error handling with specific 401/404 guidance
         - Added search configuration display showing endpoint and parameters
         - Improved search failure handling to continue with other terms
         - Better debugging output for search operations
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

# Function to get Search Service Token from main site
function Get-SearchToken {
    param(
        [string]$BaseUrl,
        [string]$TenantId,
        [string]$AccessToken
    )

    Write-Host "`nGetting Search Service Token..." -ForegroundColor Yellow

    # The search token endpoint is on the main site, not the regional search endpoint
    $authUrl = "$BaseUrl/$TenantId/search/GetSearchServiceToken"
    Write-Host "  Trying: $authUrl" -ForegroundColor Gray
    Write-Host "  Using access token: $($AccessToken.Substring(0, [Math]::Min(50, $AccessToken.Length)))..." -ForegroundColor Gray

    $headers = @{
        'Authorization' = "Bearer $AccessToken"
    }

    try {
        $response = Invoke-RestMethod -Uri $authUrl -Method Get -Headers $headers

        if ($response.Status -eq 'Success' -and $response.Message) {
            Write-Host "  Search service token retrieved successfully!" -ForegroundColor Green
            return $response.Message
        }
        elseif ($response.Message) {
            # Sometimes the response might just have a Message field directly
            Write-Host "  Search service token retrieved!" -ForegroundColor Green
            return $response.Message
        }
        else {
            Write-Host "  Unexpected response from search service token endpoint" -ForegroundColor Yellow
            Write-Host "  Response: $($response | ConvertTo-Json -Depth 2)" -ForegroundColor Gray
            Write-Host "  Will attempt to use main access token for search..." -ForegroundColor Yellow
            return $AccessToken
        }
    }
    catch {
        $statusCode = $null
        $statusDescription = $null

        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            $statusDescription = $_.Exception.Response.StatusDescription
        }

        Write-Host "  Failed to get search service token" -ForegroundColor Yellow

        if ($statusCode) {
            Write-Host "  Status Code: $statusCode - $statusDescription" -ForegroundColor Gray

            if ($statusCode -eq 404) {
                Write-Host "  NOTE: 404 error - endpoint may not exist on this Process Manager version" -ForegroundColor Yellow
            }
            elseif ($statusCode -eq 401) {
                Write-Host "  NOTE: 401 error - main access token may be invalid" -ForegroundColor Yellow
            }
        }

        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Gray
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

    # Encode the search term with quotes for literal matching
    $quotedTerm = "`"$SearchTerm`""
    $encodedTerm = [System.Web.HttpUtility]::UrlEncode($quotedTerm)

    # Build the search URL - matching the working implementation from UnpublishedProcessDocuments
    # Note: ProcessSearchFields are optional and may cause issues on some endpoints
    $searchUrl = "$SearchEndpoint/fullsearch?SearchCriteria=$encodedTerm&IncludedTypes=1&SearchMatchType=0&pageNumber=$PageNumber"

    Write-Host "    Search URL: $searchUrl" -ForegroundColor Gray
    Write-Host "    Using token: $($SearchToken.Substring(0, [Math]::Min(30, $SearchToken.Length)))..." -ForegroundColor Gray

    $headers = @{
        'Authorization' = "Bearer $SearchToken"
    }

    try {
        $response = Invoke-RestMethod -Uri $searchUrl -Method Get -Headers $headers
        return $response
    }
    catch {
        $statusCode = $null
        $statusDescription = $null

        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            $statusDescription = $_.Exception.Response.StatusDescription
        }

        Write-Host "    Search failed for term '$SearchTerm'" -ForegroundColor Yellow

        if ($statusCode) {
            Write-Host "    Status Code: $statusCode - $statusDescription" -ForegroundColor Gray

            if ($statusCode -eq 401) {
                Write-Host "    ERROR: Unauthorized (401) - Search token may be invalid or expired" -ForegroundColor Red
                Write-Host "    This usually means the search authentication token is not working" -ForegroundColor Yellow
            }
            elseif ($statusCode -eq 404) {
                Write-Host "    ERROR: Not Found (404) - Search endpoint may be incorrect" -ForegroundColor Red
            }
        }

        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Gray
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

    # Correct API endpoint: /Api/v1/Processes/ (capital A, v1, plural)
    $apiUrl = "$BaseUrl/$TenantId/Api/v1/Processes/$ProcessId"

    Write-Verbose "Getting process details: $apiUrl"

    $headers = @{
        'Authorization' = "Bearer $AccessToken"
    }

    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
        return $response
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }

        Write-Host "      ERROR: Failed to get process details for $ProcessId" -ForegroundColor Red
        if ($statusCode) {
            Write-Host "      Status Code: $statusCode" -ForegroundColor Gray
        }
        Write-Host "      URL: $apiUrl" -ForegroundColor Gray
        Write-Host "      Error: $($_.Exception.Message)" -ForegroundColor Gray
        return $null
    }
}

# Function to get process review due date
function Get-ProcessReviewDue {
    param(
        [string]$BaseUrl,
        [string]$TenantId,
        [string]$ProcessId,
        [string]$AccessToken
    )

    $apiUrl = "$BaseUrl/$TenantId/Api/v1/Processes/$ProcessId/ReviewDue"

    Write-Verbose "Getting review due date: $apiUrl"

    $headers = @{
        'Authorization' = "Bearer $AccessToken"
    }

    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
        return $response
    }
    catch {
        Write-Verbose "Could not get review due date for $ProcessId : $($_.Exception.Message)"
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

    # Get Search Service Token from main site
    $searchToken = Get-SearchToken -BaseUrl $credentials.BaseUrl -TenantId $tenantId -AccessToken $authResult.AccessToken

    if (-not $searchToken) {
        Write-Host ""
        Write-Host "ERROR: Search API authentication failed and could not fall back to main token." -ForegroundColor Red
        Write-Host ""
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }

    # Show which token we're using
    if ($searchToken -ne $authResult.AccessToken) {
        Write-Host "Using dedicated search token" -ForegroundColor Green
        Write-Host "  Token preview: $($searchToken.Substring(0, [Math]::Min(50, $searchToken.Length)))..." -ForegroundColor Gray
    }
    else {
        Write-Host "Using main access token for search (fallback)" -ForegroundColor Yellow
        Write-Host "  Token preview: $($searchToken.Substring(0, [Math]::Min(50, $searchToken.Length)))..." -ForegroundColor Gray
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
    Write-Host "Search Endpoint: $searchEndpoint" -ForegroundColor Gray
    Write-Host "Parameters: IncludedTypes=1, SearchMatchType=0 (using API defaults for field matching)" -ForegroundColor Gray
    Write-Host ""

    $allResults = @()
    $processCache = @{}

    foreach ($word in $redWords) {
        Write-Host "Searching for: '$word'" -ForegroundColor Yellow

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

                        # Get review due date
                        $reviewDue = Get-ProcessReviewDue -BaseUrl $credentials.BaseUrl -TenantId $tenantId -ProcessId $processId -AccessToken $authResult.AccessToken

                        # Cache the process
                        $processCache[$processId] = @{
                            SearchResult = $process
                            Details = $processDetails
                            ReviewDue = $reviewDue
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
            elseif ($searchResult -eq $null) {
                # Search failed - error already displayed by Search-Processes function
                Write-Host "  Skipping this search term and continuing..." -ForegroundColor Yellow
                $hasMorePages = $false
            }
            else {
                # Unexpected response format
                Write-Host "  No results found for '$word'" -ForegroundColor Gray
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
        $reviewDueData = $entry.Value.ReviewDue
        $redWordsFound = $entry.Value.RedWords -join ', '

        # Parse process title and variation name
        # If the title contains "::", the part after is the variation name
        $processTitle = $searchData.Name
        $variationName = ''

        if ($processTitle -match '::') {
            $parts = $processTitle -split '::', 2
            $processTitle = $parts[0].Trim()
            $variationName = $parts[1].Trim()
        }
        # Also check variationSetData if no :: in title
        elseif ($detailsData -and $detailsData.variationSetData) {
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

        # Get publish date from Published object (for published processes)
        $publishDate = ''
        if ($detailsData -and $detailsData.processJson -and $detailsData.processJson.Published) {
            $publishDateRaw = $detailsData.processJson.Published.PublishedDate
            if ($publishDateRaw) {
                # Parse and format the date (remove time portion)
                try {
                    $publishDate = ([DateTime]::Parse($publishDateRaw)).ToString('yyyy-MM-dd')
                }
                catch {
                    $publishDate = $publishDateRaw
                }
            }
        }

        # Get review due date
        $reviewDueDate = ''
        if ($reviewDueData -and $reviewDueData.NextReviewDate) {
            try {
                $reviewDueDate = ([DateTime]::Parse($reviewDueData.NextReviewDate)).ToString('yyyy-MM-dd')
            }
            catch {
                $reviewDueDate = $reviewDueData.NextReviewDate
            }
        }

        # Determine review status based on review due date
        $reviewStatus = 'N/A'
        if ($reviewDueDate -ne '') {
            try {
                $reviewDate = [DateTime]::Parse($reviewDueDate)
                $today = Get-Date
                if ($reviewDate -lt $today) {
                    $reviewStatus = 'Out of Date'
                }
                else {
                    $reviewStatus = 'In Date'
                }
            }
            catch {
                $reviewStatus = 'Unknown'
            }
        }

        # Create result object
        $result = [PSCustomObject]@{
            'Process Title' = $processTitle
            'Process Variation Name' = $variationName
            'Red Flag Words' = $redWordsFound
            'Process Owner' = $owner
            'Process Expert' = $expert
            'Process Group Path' = $groupPath
            'Status' = $status
            'Publish Date' = $publishDate
            'Review Due Date' = $reviewDueDate
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
