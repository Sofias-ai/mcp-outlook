# MCP Outlook Server

An MCP (Model Context Protocol) server for interacting with Microsoft Outlook through Microsoft Graph API. This server provides advanced tools for searching, managing, and cleaning emails with HTML and text content processing capabilities.

## Key Features

### 🔍 Advanced Email Search
- **Dual search methods**: OData filters + KQL (Keyword Query Language)
- **Performance optimization**: Body/No-Body variants for all search functions
- **Unified search** across multiple folders (Inbox, SentItems, Drafts)
- **Natural language searches** with KQL support ("from:john meeting today")
- **Precise filtering** with OData ("isRead eq false and importance eq 'high'")
- **Automatic pagination** for extensive results
- **Structured result formatting**

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

### 🎯 Function Selection Guide

**Choose the right function for your use case:**

| **Use Case** | **Recommended Function** | **Reason** |
|--------------|-------------------------|------------|
| **Email listing/previews** | `Search_Outlook_Emails_No_Body` or `Search_Outlook_Emails_No_Body_By_Search_Query` | Faster, reduced bandwidth |
| **Reading full email content** | `Get_Outlook_Email` | Complete email with formatted body |
| **Natural language search** | `Search_Outlook_Emails_By_Search_Query` | KQL supports content-based searches |
| **Precise filtering** | `Search_Outlook_Emails` | OData filters for exact criteria |
| **Performance-critical apps** | Functions ending with `_No_Body` | Optimized for speed |
| **Content analysis** | `Search_Outlook_Emails` or `Search_Outlook_Emails_By_Search_Query` | Full body content included |
| **Building email clients** | `Search_Outlook_Emails_No_Body` + `Get_Outlook_Email` | List view + detail view pattern |

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

#### 2b. Search_Outlook_Emails_No_Body *(Performance Optimized)*
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

#### 3. Search_Outlook_Emails_By_Search_Query *(KQL Search)*
Advanced search using Microsoft Graph's KQL (Keyword Query Language) search parameter. More powerful for natural language and content-based searches.

**Parameters:**
- `user_email` (str): User email
- `search_query` (str): KQL search query
- `top` (int, optional): Maximum number of results (default: 10)
- `folders` (List[str], optional): Folders to search (default: ["Inbox", "SentItems", "Drafts"])

**KQL Search Examples:**
```javascript
// Search by sender name or email
"from:john@example.com" or "from:John Doe"

// Search in subject and body
"meeting agenda"

// Search by recipient
"to:maria@company.com"

// Complex searches
"subject:urgent AND from:boss@company.com"

// Date-based searches
"received:today" or "received:last week"

// Attachment searches
"hasattachment:true"

// Keyword combinations
"project AND deadline NOT completed"
```

**When to use KQL vs OData:**
- **Use KQL** (`Search_Outlook_Emails_By_Search_Query`) for:
  - Natural language searches
  - Content-based searches (searching within email body)
  - Complex keyword combinations
  - Searches involving attachments
  - Date-relative searches ("today", "last week")

- **Use OData** (`Search_Outlook_Emails`) for:
  - Precise property-based filtering
  - Boolean logic on specific fields
  - Exact date/time ranges
  - Flag-based searches (isRead, importance)

#### 3b. Search_Outlook_Emails_No_Body_By_Search_Query *(KQL + Performance)*
Combines the power of KQL search with performance optimization by excluding email body content.

**Parameters:**
- `user_email` (str): User email
- `search_query` (str): KQL search query (same syntax as Search_Outlook_Emails_By_Search_Query)
- `top` (int, optional): Maximum number of results (default: 10)
- `folders` (List[str], optional): Folders to search (default: ["Inbox", "SentItems", "Drafts"])

**Best for:**
- Fast KQL-based searches when you only need email metadata
- Building search result previews with KQL capabilities
- Performance-critical applications using natural language search

**Example response for all search functions:**
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
      // Note: 'body' field only included in full search functions
    }
  ]
}
```

#### 4. Create_Outlook_Draft_Email
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

#### 5. Update_Outlook_Draft_Email
Updates an existing draft, including adding new attachments.

**Parameters:**
- `message_id` (str): ID of the draft to update
- `user_email` (str): User email
- All optional parameters from `Create_Outlook_Draft_Email`

#### 6. Delete_Outlook_Email
Deletes an email by its ID.

**Parameters:**
- `message_id` (str): ID of the message to delete
- `user_email` (str): User email

## 🆕 New in 0.1.11

### Reply Drafts with Preserved Conversation History
- The `create_draft_email_tool` function now supports a `reply_to_id` parameter.
- When replying to a message, the draft will automatically include the full content (including formatting and conversation history) of the original message, just like Outlook does.
- No extra formatting is applied: the original HTML or plain text is preserved for maximum fidelity.
- This makes replies and conversation threads in drafts visually identical to those in the Outlook client.

#### Example usage
```python
create_draft_email_tool(
    subject="Re: Project Update",
    body="Thank you for the update!",
    to_recipients=["user@example.com"],
    user_email="me@example.com",
    reply_to_id="AAMkAGI2..."
)
```
This will create a draft that includes your reply and the full original message content, preserving all formatting and previous conversation history.

## 💡 Practical Examples

### Example 1: Building an Email Client Interface
```python
# Step 1: Get email list for preview (fast)
emails = search_emails_no_body_tool(
    user_email="user@company.com",
    query_filter="isRead eq false",
    top=20
)

# Step 2: When user selects an email, get full content
full_email = get_email_tool(
    message_id=selected_email_id,
    user_email="user@company.com"
)
```

### Example 2: Smart Content Search
```python
# Natural language search using KQL
meetings = search_emails_by_search_query_tool(
    user_email="user@company.com", 
    search_query="meeting AND (today OR tomorrow)",
    top=10
)

# Precise filtering using OData
urgent_unread = search_emails_tool(
    user_email="user@company.com",
    query_filter="isRead eq false and importance eq 'high'",
    top=5
)
```

### Example 3: Performance-Optimized Searches
```python
# Fast search for email previews without body processing
preview_results = search_emails_no_body_by_search_query_tool(
    user_email="user@company.com",
    search_query="from:boss@company.com project",
    top=50
)
```

### Example 4: Draft Management
```python
# Create draft with attachments
draft = create_draft_email_tool(
    subject="Weekly Report",
    body="<h1>Weekly Status</h1><p>Please find attached...</p>",
    to_recipients=["team@company.com"],
    user_email="user@company.com",
    category="Work",
    file_paths=["C:/reports/weekly.pdf"]
)

# Update the draft
update_draft_email_tool(
    message_id=draft['id'],
    user_email="user@company.com",
    subject="Weekly Report - Updated",
    cc_recipients=["manager@company.com"]
)
```

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