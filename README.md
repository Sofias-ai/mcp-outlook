# Outlook MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A lightweight MCP Server for seamless integration with Microsoft Outlook, enabling MCP clients to interact with emails for specific users. Developed by [sofias tech](https://github.com/sofias/).

## Features

This server provides a clean interface to Outlook email resources through the Model Context Protocol (MCP), with operations for reading, searching, creating, updating, and deleting emails.

### Tools

The server implements the following tools:

- `Get_Outlook_Email`: Retrieves a specific email by its unique ID.
- `Search_Outlook_Emails`: Searches emails using OData filter syntax (e.g., "subject eq 'Update'", "isRead eq false") across multiple folders.
- `Download_Outlook_Emails_By_Date`: Downloads emails within a specific date range using ISO 8601 format (YYYY-MM-DDTHH:MM:SSZ).
- `Create_Outlook_Draft_Email`: Creates a new draft email with the specified subject, body, and recipients.
- `Update_Outlook_Draft_Email`: Updates an existing draft email specified by its ID.
- `Delete_Outlook_Email`: Deletes an email by its ID (moves it to the Deleted Items folder).

All tools require a `user_email` parameter to specify which mailbox to access.

## Architecture

The server is built with efficiency and clarity in mind:

- Utilizes the Microsoft Graph API via the `office365-rest-python-client` library
- Modular design with clear separation between:
  - `resources.py`: Core API functionality for interacting with Outlook
  - `tools.py`: MCP tool definitions that wrap resource functions
  - `common.py`: Shared components and configuration
  - `server.py`: Main entry point that initializes and runs the MCP server
- Uses environment variables for secure configuration
- Supports accessing multiple user mailboxes through application permissions

## Setup

1. Register an app in Azure AD
2. Grant the necessary Microsoft Graph API permissions at the application level (not delegated):
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `User.Read`
3. Obtain the client ID, tenant ID, and create a client secret
4. Have an admin grant consent for these permissions at the organization level

## Environment Variables

Create a `.env` file with the following variables:

- `ID_CLIENT`: Your Azure AD application client ID.
- `APP_SECRET`: Your Azure AD application client secret.
- `TENANT_ID`: Your Microsoft Directory (tenant) ID.

Note: The user email is now passed as a parameter to each tool rather than being defined as an environment variable.

## Quickstart

### Installation

Install the package in editable mode for development:

```bash
pip install -e .
```

Or install from PyPI once published:

```bash
pip install mcp-outlook
```

Using uv:

```bash
uv pip install mcp-outlook
```

### Claude Desktop Integration

To integrate with Claude Desktop, update the configuration file:

On Windows: `%APPDATA%/Claude/claude_desktop_config.json`
On macOS: `~/Library/Application\ Support/Claude/claude_desktop_config.json`

#### Standard Integration

```json
"mcpServers": {
  "outlook": {
    "command": "mcp-outlook",
    "env": {
      "ID_CLIENT": "your-app-id",
      "APP_SECRET": "your-app-secret",
      "TENANT_ID": "your-tenant-id"
    }
  }
}
```

#### Using uvx

```json
"mcpServers": {
  "outlook": {
    "command": "uvx",
    "args": [
      "mcp-outlook"
    ],
    "env": {
      "ID_CLIENT": "your-app-id",
      "APP_SECRET": "your-app-secret",
      "TENANT_ID": "your-tenant-id"
    }
  }
}
```

## Development

### Requirements

- Python 3.10+
- Dependencies listed in `pyproject.toml`

### Local Development

1. Clone the repository.
2. Create a virtual environment:

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```
3. Install development dependencies:

   ```bash
   pip install -e .
   ```
4. Create a `.env` file in the project root with your Azure AD app credentials:

   ```dotenv
   ID_CLIENT=your-app-id
   APP_SECRET=your-app-secret
   TENANT_ID=your-tenant-id
   ```
5. Run the server:

   ```bash
   python -m mcp_outlook_server
   ```

   *Note: Ensure the module name matches your main execution script if different.*

### Debugging

For debugging the MCP server, you can use the [MCP Inspector](https://github.com/modelcontextprotocol/inspector):

```bash
npx @modelcontextprotocol/inspector -- python -m mcp_outlook_server
```

## Example Usage

When using the tools through Claude or another MCP client, you'll need to specify the user email for each operation:

- "Get emails from last week for user@example.com"
- "Search for emails with 'Project Update' in the subject for john.doe@company.com"
- "Create a draft email for sarah@example.com to send to the team"
- "Delete the email with ID ABC123 from user@example.com's inbox"

## License

This project is licensed under the MIT License - see the LICENSE file for details.

Copyright (c) 2024 sofias tech
