# MCP Outlook Server

An MCP (Model Context Protocol) server for interacting with Microsoft Outlook through Microsoft Graph API. This server provides advanced tools for searching, managing, and cleaning emails with HTML and text content processing capabilities.

## Key Features

### 🔍 Advanced Email Search
- Unified search across multiple folders (Inbox, SentItems, Drafts)
- Full support for OData filters
- Automatic pagination for extensive results
- Structured result formatting
- **NEW**: Performance-optimized search without email body content

### 📧 Email Management
- Draft creation with local attachments
- Updating existing drafts
- Email deletion
- Support for categories and multiple recipients (TO, CC, BCC)

### 🧹 Advanced Content Cleaning
- Intelligent HTML to plain text conversion
- Automatic noise and disclaimer removal
- Character and format normalization
- Boilerplate detection and truncation

### 📊 Structured Formatting
- Consistent dictionary format output
- Automatic metadata extraction (sender, date, subject, etc.)
- Automatic content summaries
- Robust error handling

## Installation

### Prerequisites
- Python 3.8+
- Azure AD application registration with Microsoft Graph permissions
- Configured environment variables

### Dependencies
```bash
pip install mcp fastmcp office365-rest-python-client python-dotenv beautifulsoup4 html2text
```

### Azure AD Configuration

1. Register a new application in Azure AD
2. Configure the following Microsoft Graph permissions:
   - `Mail.Read` - Read user emails
   - `Mail.ReadWrite` - Create, update and delete emails
   - `Mail.Send` - Send emails (if needed)
   - `User.Read` - Read basic user profile

3. Generate a client secret

### Environment Variables

Create a `.env` file in the project root:

```env
ID_CLIENT=your_azure_client_id
APP_SECRET=your_azure_client_secret
TENANT_ID=your_azure_tenant_id
```

## Usage

### Starting the Server
```bash
python -m mcp_outlook_server.server
```

### Available Tools

#### 1. Get_Outlook_Email
Retrieves a specific email by its unique ID.

**Parameters:**
- `message_id` (str): Unique message ID
- `user_email` (str): Email of the mailbox owner

**Example response:**
```json
{
  "success": true,
  "data": {
    "sender": "John Doe <john@example.com>",
    "date": "2024-01-15 10:30:00",
    "cc": ["maria@example.com"],
    "subject": "Project meeting",
    "summary": "Meeting confirmation to discuss project progress...",
    "body": "Complete email content clean and formatted",
    "id": "AAMkAGE1M2IyNGNmLWI4MjktNDUyZi1iMzA4LTViNDI3NzhlOGM2NgBGAAAAAADUuTiuQqVlSKDGAz"
  }
}
```

#### 2. Search_Outlook_Emails
Advanced email search with OData filter support.

**Parameters:**
- `user_email` (str): User email
- `query_filter` (str, optional): OData filter
- `top` (int, optional): Maximum number of results (default: 10)
- `folders` (List[str], optional): Folders to search (default: ["Inbox", "SentItems", "Drafts"])

**OData filter examples:**
```javascript
// Unread emails
"isRead eq false"

// Emails from specific sender
"from/emailAddress/address eq 'user@example.com'"

// Emails with specific subject
"subject eq 'Monthly report'"

// Emails received in last 7 days
"receivedDateTime ge 2024-01-01T00:00:00Z"

// Combined filters
"isRead eq false and importance eq 'high'"
```

#### 2b. Search_Outlook_Emails_No_Body *(New)*
Performance-optimized email search that excludes email body content for faster processing and reduced data transfer.

**Parameters:**
- `user_email` (str): User email
- `query_filter` (str, optional): OData filter (same syntax as Search_Outlook_Emails)
- `top` (int, optional): Maximum number of results (default: 10)
- `folders` (List[str], optional): Folders to search (default: ["Inbox", "SentItems", "Drafts"])

**Key Benefits:**
- **Faster performance**: Excludes body content processing
- **Reduced bandwidth**: Smaller response payload
- **Uses native Microsoft Graph `bodyPreview` field** for the summary, which is more efficient and consistent than manual summary generation
- **Ideal for listings**: Perfect for email previews and list views
- **Same filtering**: Full OData filter support maintained

**Use Cases:**
- Quick email listing and previews
- Performance-critical searches when body content is not needed
- Bandwidth-constrained environments
- Building email selection interfaces

**Example response:**
```json
{
  "success": true,
  "data": [
    {
      "sender": "John Doe <john@example.com>",
      "date": "2024-01-15 10:30:00",
      "cc": ["maria@example.com"],
      "subject": "Project meeting",
      "summary": "Hi team, I wanted to schedule our weekly project review for Thursday at 2pm. Please confirm your availability and we'll send out the meeting invite. The agenda will cover...",
      "id": "AAMkAGE1M2IyNGNmLWI4MjktNDUyZi1iMzA4LTViNDI3NzhlOGM2NgBGAAAAAADUuTiuQqVlSKDGAz"
      // Note: 'body' and 'cuerpo' fields are excluded, and 'summary' uses Microsoft Graph's bodyPreview
    }
  ]
}
```

#### 3. Create_Outlook_Draft_Email
Creates a new draft email with support for attachments and categories.

**Parameters:**
- `subject` (str): Email subject
- `body` (str): Message body
- `to_recipients` (List[str]): List of primary recipients
- `user_email` (str): User email
- `cc_recipients` (List[str], optional): CC recipients
- `bcc_recipients` (List[str], optional): BCC recipients
- `body_type` (str, optional): Body type ("HTML" or "Text")
- `category` (str, optional): Email category
- `file_paths` (List[str], optional): File paths to attach

**Example:**
```python
create_draft_email_tool(
    subject="Project report",
    body="<p>Please find the requested report attached.</p>",
    to_recipients=["client@example.com"],
    cc_recipients=["supervisor@company.com"],
    user_email="my.email@company.com",
    category="Work",
    file_paths=["C:/documents/report.pdf", "C:/documents/data.xlsx"]
)
```

#### 4. Update_Outlook_Draft_Email
Updates an existing draft, including adding new attachments.

**Parameters:**
- `message_id` (str): ID of the draft to update
- `user_email` (str): User email
- All optional parameters from `Create_Outlook_Draft_Email`

#### 5. Delete_Outlook_Email
Deletes an email by its ID.

**Parameters:**
- `message_id` (str): ID of the message to delete
- `user_email` (str): User email

## Technical Features

### Content Cleaning System

The server includes an advanced cleaning system that:

1. **Converts HTML to plain text** using html2text with optimized configuration
2. **Removes unwanted elements**: scripts, styles, metadata, etc.
3. **Normalizes text**: special characters, spaces, line breaks
4. **Detects and truncates disclaimers**: identifies legal/boilerplate text patterns
5. **Generates automatic summaries**: extracts first 150 relevant characters

### Robust Error Handling

- Exception handling decorators on all critical operations
- Fallback strategies for content processing
- Detailed logging for debugging
- Input parameter validation

### Performance Optimizations

- Automatic pagination for extensive searches
- Page limits to avoid timeouts
- Result caching when appropriate
- Parallel content processing when possible

## Project Structure

```
mcp-outlook/
├── src/
│   └── mcp_outlook_server/
│       ├── __init__.py
│       ├── server.py          # Server entry point
│       ├── common.py          # Configuration and common utilities
│       ├── tools.py           # MCP tool definitions
│       ├── resources.py       # Resource access logic
│       ├── format_utils.py    # Format and extraction utilities
│       └── clean_utils.py     # Content cleaning system
├── .env                       # Environment variables (don't include in git)
├── requirements.txt
└── README.md
```

## Logging

The server generates detailed logs in:
- File: `mcp_outlook.log`
- Console: Standard output

Logging configuration in `common.py` with INFO level by default.

## Troubleshooting

### Authentication Error
- Verify environment variables are correctly configured
- Confirm Azure AD application has necessary permissions
- Check that tenant ID is correct

### Search Errors
- Validate OData filter syntax
- Verify user has access to specified folders
- Check logs for specific Graph API errors

### Attachment Issues
- Confirm file paths exist and are accessible
- Verify files don't exceed Outlook size limits
- Check file read permissions

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-functionality`)
3. Commit changes (`git commit -am 'Add new functionality'`)
4. Push to branch (`git push origin feature/new-functionality`)
5. Create a Pull Request

## License

MIT License

Copyright (c) 2024

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Support

To report bugs or request new features, please open an issue in the repository.