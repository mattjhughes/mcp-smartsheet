# Smartsheet MCP Server

A Model Context Protocol (MCP) server for integrating with the Smartsheet API. This server allows AI assistants to interact with Smartsheet data through a standardized interface.

## Features

- **List Available Sheets**: View all accessible Smartsheet spreadsheets
- **Read Sheet Data**: Access the contents of specific sheets
- **Add Rows**: Add new rows to spreadsheets
- **Update Rows**: Update existing row data
- **List Columns**: Get column information for a sheet
- **Search**: Search across all sheets for specific text

## Prerequisites

- Node.js (v16 or later)
- npm or yarn
- A Smartsheet API token (generated from Smartsheet UI)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/smartsheet-mcp-server.git
   cd smartsheet-mcp-server
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Build the project:
   ```bash
   npm run build
   ```

## Usage

### Running the server

```bash
SMARTSHEET_API_TOKEN=your_api_token npm start
```

### MCP Configuration

An example MCP configuration is provided in `example-mcp-config.json`:

```json
{
  "mcpServers": {
    "smartsheet": {
      "command": "node",
      "args": ["./dist/index.js"],
      "env": {
        "SMARTSHEET_API_TOKEN": "YOUR_API_KEY_HERE"
      }
    }
  }
}
```

This configuration can be used with any MCP-compatible client.

### Running the server with Claude

An example configuration file for Claude is provided in `example-claude-config.json`:

```json
{
  "mcpServers": {
    "smartsheet": {
      "command": "node",
      "args": ["/path/to/mcp-smartsheet/dist/index.js"],
      "env": {
        "SMARTSHEET_API_TOKEN": "YOUR_SMARTSHEET_API_TOKEN"
      }
    }
  }
}
```

1. Copy `example-claude-config.json` to a location of your choice
2. Replace `/path/to/mcp-smartsheet` with the actual path to your installation
3. Replace `YOUR_SMARTSHEET_API_TOKEN` with your actual Smartsheet API token
4. Configure Claude to use this configuration file

## How to Get a Smartsheet API Token

1. Log in to your Smartsheet account
2. Go to Account > Personal Settings > API Access
3. Click "Generate new access token"
4. Copy the token (you won't be able to see it again)

## Available Tools

The server provides the following tools for AI assistants:

- **listSheets**: Get a list of all available sheets
- **getSheet**: Get the complete content of a sheet including rows and cells
- **addRow**: Add a new row to a sheet
- **updateRow**: Update an existing row in a sheet
- **listColumns**: Get column definitions for a sheet
- **searchSheets**: Search across sheets for specific text

## Security Considerations

- Never commit your Smartsheet API token to the repository
- Use environment variables or secure configuration services to manage API tokens
- Consider setting up proper access controls on your Smartsheet account
- The server runs locally on your machine

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.