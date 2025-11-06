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
    Version: 1.0
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

# Function to get credentials
function Get-ProcessManagerCredentials {
    Write-Host "`n=== Process Manager Red Word Search Tool ===" -ForegroundColor Cyan
    Write-Host ""

    # Get URL
    $url = Read-Host "Enter your Process Manager URL (e.g., https://demo.promapp.com)"

    # Validate and clean URL
    if ($url -notmatch '^https?://') {
        $url = "https://$url"
    }
    $url = $url.TrimEnd('/')

    # Get credentials
    Write-Host ""
    $username = Read-Host "Enter your username"
    $securePassword = Read-Host "Enter your password" -AsSecureString
    $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))

    return @{
        Url = $url
        Username = $username
        Password = $password
    }
}

# Function to determine search endpoint from URL
function Get-SearchEndpoint {
    param([string]$BaseUrl)

    $uri = [System.Uri]$BaseUrl
    $host = $uri.Host

    if ($RegionalEndpoints.ContainsKey($host)) {
        return $RegionalEndpoints[$host]
    }

    Write-Warning "Unknown region for host: $host. Using Demo region endpoint."
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
        [string]$Password
    )

    Write-Host "`nAuthenticating to Process Manager..." -ForegroundColor Yellow

    # First, we need to get the tenant ID by making an initial request
    # Try to access the main page to get redirected or find tenant
    try {
        $response = Invoke-WebRequest -Uri $BaseUrl -UseBasicParsing -MaximumRedirection 0 -ErrorAction SilentlyContinue
    }
    catch {
        # This is expected, we're just trying to find the tenant
    }

    # Try common authentication endpoint pattern
    # The tenant ID is usually part of the URL path
    $authUrl = "$BaseUrl/oauth2/token"

    $body = @{
        grant_type = 'password'
        username = $Username
        password = $Password
    }

    try {
        $response = Invoke-RestMethod -Uri $authUrl -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded'

        if ($response.access_token) {
            Write-Host "Authentication successful!" -ForegroundColor Green
            return @{
                AccessToken = $response.access_token
                TokenType = $response.token_type
                ExpiresIn = $response.expires_in
            }
        }
    }
    catch {
        Write-Error "Authentication failed: $($_.Exception.Message)"

        # Try to extract tenant from the base URL by making a request to the site
        try {
            $mainPage = Invoke-WebRequest -Uri $BaseUrl -UseBasicParsing
            if ($mainPage.Content -match '/([a-f0-9]{32})/') {
                $tenantId = $matches[1]
                Write-Host "Found tenant ID: $tenantId" -ForegroundColor Cyan

                $authUrl = "$BaseUrl/$tenantId/oauth2/token"
                $response = Invoke-RestMethod -Uri $authUrl -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded'

                if ($response.access_token) {
                    Write-Host "Authentication successful!" -ForegroundColor Green
                    return @{
                        AccessToken = $response.access_token
                        TokenType = $response.token_type
                        ExpiresIn = $response.expires_in
                        TenantId = $tenantId
                    }
                }
            }
        }
        catch {
            Write-Error "Could not authenticate. Please verify your credentials and URL."
            return $null
        }
    }

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

    Write-Host "Authenticating to Search API..." -ForegroundColor Yellow

    $authUrl = "$SearchEndpoint/api/authentication/$TenantId/$UserId"

    $headers = @{
        'Authorization' = "Bearer $AccessToken"
        'Content-Type' = 'application/json'
    }

    try {
        $response = Invoke-RestMethod -Uri $authUrl -Method Post -Headers $headers

        if ($response.Status -eq 'Success' -and $response.Message) {
            Write-Host "Search API authentication successful!" -ForegroundColor Green
            return $response.Message
        }
    }
    catch {
        Write-Warning "Search API authentication failed: $($_.Exception.Message)"
        Write-Host "Will attempt to use main access token for search..." -ForegroundColor Yellow
        return $AccessToken
    }

    return $null
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
    $searchEndpoint = Get-SearchEndpoint -BaseUrl $credentials.Url
    Write-Host "`nUsing search endpoint: $searchEndpoint" -ForegroundColor Cyan

    # Authenticate to Process Manager
    $authResult = Get-ProcessManagerToken -BaseUrl $credentials.Url -Username $credentials.Username -Password $credentials.Password

    if (-not $authResult -or -not $authResult.AccessToken) {
        Write-Error "Authentication failed. Exiting."
        return
    }

    # Extract tenant ID from token
    $tenantId = Get-TenantId -BaseUrl $credentials.Url -AccessToken $authResult.AccessToken

    if (-not $tenantId) {
        # Try to extract from auth result if available
        if ($authResult.TenantId) {
            $tenantId = $authResult.TenantId
        }
        else {
            Write-Error "Could not determine tenant ID. Exiting."
            return
        }
    }

    Write-Host "Tenant ID: $tenantId" -ForegroundColor Cyan

    # Get user ID from token
    $tokenParts = $authResult.AccessToken.Split('.')
    if ($tokenParts.Count -ge 2) {
        $payload = $tokenParts[1]
        $padding = (4 - ($payload.Length % 4)) % 4
        $payload += "=" * $padding
        $decodedBytes = [System.Convert]::FromBase64String($payload)
        $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
        $tokenData = $decodedJson | ConvertFrom-Json
        $userId = $tokenData.UserId
    }

    # Authenticate to Search API
    $searchToken = Get-SearchToken -SearchEndpoint $searchEndpoint -TenantId $tenantId -UserId $userId -AccessToken $authResult.AccessToken

    if (-not $searchToken) {
        Write-Error "Search API authentication failed. Exiting."
        return
    }

    # Get red flag words
    $redWords = Get-RedWords

    if ($redWords.Count -eq 0) {
        Write-Error "No red flag words specified. Exiting."
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
                        $processDetails = Get-ProcessDetails -BaseUrl $credentials.Url -TenantId $tenantId -ProcessId $processId -AccessToken $authResult.AccessToken

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
}

# Run the script
Main
