# Outlook MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A lightweight MCP Server for seamless integration with Microsoft Outlook, enabling MCP clients to interact with emails for a specific user. Developed by [sofias tech](https://github.com/sofias/).

## Features

This server provides a clean interface to Outlook email resources through the Model Context Protocol (MCP), with operations for reading, searching, creating, updating, and deleting emails.

### Tools

The server implements the following tools:

-   `Get_Outlook_Email`: Retrieves a specific email by its unique ID.
-   `Search_Outlook_Emails`: Searches emails using OData filter syntax (e.g., "subject eq 'Update'", "isRead eq false"). Returns a list of emails.
-   `Download_Outlook_Emails_By_Date`: Downloads emails received within a specific date range (ISO 8601 format). Returns a list of emails.
-   `Create_Outlook_Draft_Email`: Creates a new draft email with the specified subject, body, and recipients.
-   `Update_Outlook_Draft_Email`: Updates an existing draft email specified by its ID. Only provided fields are updated.
-   `Delete_Outlook_Email`: Deletes an email by its ID (moves it to the Deleted Items folder).

## Architecture

The server is built with resource efficiency in mind:

-   Utilizes the Microsoft Graph API via the `office365-rest-python-client` library.
-   Error handling through decorators for cleaner tool implementation.
-   Clear separation between resource functions (email operations) and tool definitions.
-   Uses environment variables for secure configuration.

## Setup

1.  Register an app in Azure AD.
2.  Grant the necessary Microsoft Graph API permissions (e.g., `Mail.ReadWrite`, `Mail.Send`, `User.Read` - Application permissions).
3.  Obtain the Application (client) ID, Directory (tenant) ID, and create a client secret for the registered app.
4.  Ensure the target user mailbox exists.

## Environment Variables

The server requires these environment variables, typically stored in a `.env` file:

-   `OUTLOOK_ID_APP`: Your Azure AD application client ID.
-   `OUTLOOK_ID_APP_SECRET`: Your Azure AD application client secret.
-   `OUTLOOK_TENANT_ID`: Your Microsoft Directory (tenant) ID.
-   `OUTLOOK_USER_EMAIL`: The email address of the user whose mailbox will be accessed.

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
      "OUTLOOK_ID_APP": "your-app-id",
      "OUTLOOK_ID_APP_SECRET": "your-app-secret",
      "OUTLOOK_TENANT_ID": "your-tenant-id",
      "OUTLOOK_USER_EMAIL": "user@yourdomain.com"
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
      "OUTLOOK_ID_APP": "your-app-id",
      "OUTLOOK_ID_APP_SECRET": "your-app-secret",
      "OUTLOOK_TENANT_ID": "your-tenant-id",
      "OUTLOOK_USER_EMAIL": "user@yourdomain.com"
    }
  }
}
```

## Development

### Requirements

-   Python 3.10+
-   Dependencies listed in `pyproject.toml`

### Local Development

1.  Clone the repository.
2.  Create a virtual environment:
    ```bash
    python -m venv .venv
    source .venv/bin/activate  # On Windows: .venv\Scripts\activate
    ```
3.  Install development dependencies:
    ```bash
    pip install -e .
    ```
4.  Create a `.env` file in the project root with your Azure AD app credentials and target user email:
    ```dotenv
    OUTLOOK_ID_APP=your-app-id
    OUTLOOK_ID_APP_SECRET=your-app-secret
    OUTLOOK_TENANT_ID=your-tenant-id
    OUTLOOK_USER_EMAIL=user@yourdomain.com
    ```
5.  Run the server:
    ```bash
    python -m mcp_outlook_server
    ```
    *Note: Ensure the module name matches your main execution script if different.*

### Debugging

For debugging the MCP server, you can use the [MCP Inspector](https://github.com/modelcontextprotocol/inspector):

```bash
npx @modelcontextprotocol/inspector -- python -m mcp_outlook_server
```

## License

This project is licensed under the MIT License - see the LICENSE file for details (assuming you will add one).

Copyright (c) 2024 sofias tech
