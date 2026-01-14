# Google Sheets & Forms MCP Server

A simple MCP server that provides tools, resources, and prompts for working with Google Sheets and Google Forms.

## Features

### Tools
- `list_spreadsheets` - List your Google Spreadsheets
- `read_sheet` - Read data from a Google Sheet
- `write_sheet` - Write data to a Google Sheet
- `append_sheet` - Append data to a Google Sheet
- `create_spreadsheet` - Create a new spreadsheet
- `create_form` - Create a new Google Form

### Resources
- `sheet://{spreadsheet_id}/{range_name}` - Access sheet data as a resource

### Prompts
- `analyze_sheet_data` - Template for analyzing sheet data
- `create_report_template` - Template for creating report spreadsheets
- `form_to_sheet` - Template for form-to-sheet workflow

## Setup

### 1. Create virtual environment
```bash
cd gsheets-mcp-server
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Get Google API Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing one
3. Enable the following APIs:
   - Google Sheets API
   - Google Forms API
   - Google Drive API
4. Go to "Credentials" → "Create Credentials" → "OAuth client ID"
5. Choose "Desktop app" as application type
6. Download the credentials JSON file
7. Save it as `credentials.json` in this directory

### 4. Test the server with MCP Inspector
```bash
npx @modelcontextprotocol/inspector python gsheets_server.py
```

## Claude Desktop Configuration

Add this to your Claude Desktop config file:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "gsheets": {
      "command": "python",
      "args": [
        "/FULL/PATH/TO/gsheets-mcp-server/gsheets_server.py"
      ],
      "env": {
        "GOOGLE_CREDENTIALS_PATH": "/FULL/PATH/TO/gsheets-mcp-server/credentials.json"
      }
    }
  }
}
```

**Important**: Replace `/FULL/PATH/TO/` with the actual absolute path to this directory.

## First Run

On first run, the server will open a browser window for Google authentication. After you authorize the app, a `token.pickle` file will be created to store your credentials for future use.

## Example Usage in Claude Desktop

Once configured, you can ask Claude:

- "List my spreadsheets"
- "Read the data from spreadsheet ID 'abc123' in range 'Sheet1!A1:C10'"
- "Create a new spreadsheet called 'Q1 Report'"
- "Append this data to my sheet: [['Name', 'Email'], ['John', 'john@example.com']]"
- "Use the analyze_sheet_data prompt"
- "Show me the sheet resource for spreadsheet 'abc123'"

## Getting Spreadsheet IDs

The spreadsheet ID is in the URL:
```
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit
```

## Troubleshooting

- **Authentication fails**: Delete `token.pickle` and run again
- **Permission denied**: Check that you've enabled the APIs in Google Cloud Console
- **Module not found**: Make sure virtual environment is activated and dependencies are installed