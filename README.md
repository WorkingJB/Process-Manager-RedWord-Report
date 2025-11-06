# Process Manager Red Word Report Tool

A PowerShell script that searches Process Manager for processes containing "red flag" words and generates a detailed CSV report.

## Features

- Interactive authentication with Process Manager
- Automatic regional search endpoint detection
- Support for multiple red flag words
- Comprehensive process details including:
  - Process Title
  - Process Variation name (if applicable)
  - Red flag words identified
  - Process Owner / Expert
  - Process group path
  - Status (Published, Unpublished, In Progress)
  - Publish date (if published)
  - Review status
  - Process URL
- CSV export with timestamp
- Progress tracking and summary statistics

## Requirements

- PowerShell 5.1 or later
- Network access to Process Manager
- Valid Process Manager credentials

## Regional Endpoints

The script automatically detects and uses the correct search endpoint based on your Process Manager URL:

| Region    | Base URL                  | Search Endpoint                    |
|-----------|---------------------------|------------------------------------|
| Demo      | https://demo.promapp.com  | https://dmo-wus-sch.promapp.io    |
| US        | https://us.promapp.com    | https://prd-wus-sch.promapp.io    |
| Canada    | https://ca.promapp.com    | https://prd-cac-sch.promapp.io    |
| Europe    | https://eu.promapp.com    | https://prd-neu-sch.promapp.io    |
| Australia | https://au.promapp.com    | https://prd-aus-sch.promapp.io    |

## Usage

### Basic Usage

1. Run the script:
   ```powershell
   .\Search-RedWordProcesses.ps1
   ```

2. Follow the prompts:
   - Enter your Process Manager URL (e.g., `https://demo.promapp.com`)
   - Enter your username
   - Enter your password
   - Choose to enter red flag words manually or load from a file

3. The script will:
   - Authenticate to Process Manager
   - Search for all processes containing the specified words
   - Retrieve detailed information for each process
   - Export results to a timestamped CSV file

### Red Flag Words Input

#### Option 1: Manual Entry
Enter red flag words as comma-separated values when prompted:
```
Enter red flag words (comma-separated):
classified, confidential, secret, restricted
```

#### Option 2: File Input
Create a text file with one red flag word per line (see `sample-redwords.txt`):
```
classified
confidential
secret
restricted
sensitive
proprietary
```

Then select option 2 when prompted and provide the file path.

## Output

The script generates a CSV file named `RedWordProcesses_YYYYMMDD_HHMMSS.csv` with the following columns:

- **Process Title**: The name of the process
- **Process Variation Name**: The variation name if the process has variations
- **Red Flag Words**: Comma-separated list of red flag words found in this process
- **Process Owner**: The person who owns the process
- **Process Expert**: The subject matter expert for the process
- **Process Group Path**: The full group hierarchy path
- **Status**: Published, Unpublished, or In Progress
- **Publish Date**: The date the process was published (if applicable)
- **Review Status**: Whether the process is in date or out of date
- **Process URL**: Direct link to the process in Process Manager
- **Process ID**: Unique identifier for the process

## Example Output

```
Process Title,Process Variation Name,Red Flag Words,Process Owner,Process Expert,Process Group Path,Status,Publish Date,Review Status,Process URL,Process ID
"Security Protocol",,classified,John Doe,Jane Smith,"Company Ltd > IT > Security",Published,2024-01-15,In Date,https://demo.promapp.com/...,abc123...
"Data Handling",,confidential,Jane Smith,John Doe,"Company Ltd > Compliance",In Progress,,,https://demo.promapp.com/...,def456...
```

## API Examples

This repository includes example API responses for reference:

- `ExampleAuthOutput.json` - Main authentication response
- `ExampleSearchAuthOutput.json` - Search API authentication response
- `ExampleSearchOutput.json` - Search results response
- `ExampleGetProcess.json` - Process details response
- `ExampleSpec.json` - OpenAPI specification

## Troubleshooting

### Authentication Issues

If authentication fails:
1. Verify your URL is correct (include `https://`)
2. Check your username and password
3. Ensure you have network access to Process Manager
4. Check if your organization uses SSO (this script uses basic authentication)

### Search Issues

If searches return no results:
1. Verify the red flag words exist in your processes
2. Check that you have permission to view the processes
3. Try searching for a common word first to test

### API Rate Limiting

If you're searching for many processes:
1. The script includes built-in paging support
2. Consider running during off-peak hours
3. Split large red word lists into smaller batches

## Advanced Usage

### Verbose Output

For detailed logging, run with verbose output:
```powershell
.\Search-RedWordProcesses.ps1 -Verbose
```

### Modifying the Script

The script is organized into functions for easy customization:
- `Get-ProcessManagerCredentials`: Modify credential collection
- `Get-SearchEndpoint`: Add custom regional endpoints
- `Search-Processes`: Adjust search parameters
- `Get-ProcessDetails`: Modify detail retrieval
- `Main`: Change the overall workflow

## Security Notes

- Credentials are collected securely using `Read-Host -AsSecureString`
- Passwords are not stored or logged
- Authentication tokens are kept in memory only
- Consider using a credential manager for frequent use

## License

This tool is provided as-is for use with Process Manager.

## Contributing

Issues and pull requests are welcome at: https://github.com/WorkingJB/Process-Manager-RedWord-Report

## Related Projects

- [Unpublished Process Documents](https://github.com/WorkingJB/UnpublishedProcessDocuments) - Related Process Manager tools

## Version History

### Version 1.0
- Initial release
- Support for all major regions
- CSV export functionality
- Interactive credential collection
- Automatic tenant detection
- Paging support for large result sets
