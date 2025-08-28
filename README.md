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
- Optional lightweight mode to skip deep HTML cleaning (`clean_body=False`)

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
CLIENT_ID=your_azure_client_id
CLIENT_SECRET=your_azure_client_secret
TENANT_ID=your_azure_tenant_id
# Optional: comma separated folders to search by default
OUTLOOK_DEFAULT_FOLDERS=Inbox,SentItems,Drafts
```

**Note:** Make sure to use exactly these variable names (`CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`) as they are required by the server.

## Usage

### Starting the Server
```bash
python -m mcp_outlook_server.server
```

### 🎯 Function Selection Guide

**Choose the right function for your use case:**

| **Use Case** | **Recommended Function** | **Reason** |
|--------------|-------------------------|------------|
| **Email listing/previews** | `Search_Outlook_Emails_No_Body` or `Search_Outlook_Emails_No_Body_By_Search_Query` | Faster, reduced bandwidth, optimized for UI |
| **Reading full email content** | `Get_Outlook_Email` | Complete email with formatted body |
| **Natural language search** | `Search_Outlook_Emails_By_Search_Query` | KQL supports content-based searches |
| **Precise filtering** | `Search_Outlook_Emails` | OData filters for exact criteria |
| **Performance-critical apps** | Functions ending with `_No_Body` | Optimized for speed, uses native Graph bodyPreview |
| **Content analysis** | `Search_Outlook_Emails` or `Search_Outlook_Emails_By_Search_Query` | Full body content included |
| **Building email clients** | `Search_Outlook_Emails_No_Body` + `Get_Outlook_Email` | List view + detail view pattern |
| **Complex business logic** | OData functions (`Search_Outlook_Emails*`) | Precise boolean logic and date ranges |
| **User-friendly search** | KQL functions (`*_By_Search_Query`) | Natural language and keyword searches |

### 🔍 Search Method Comparison

| **Feature** | **OData ($filter)** | **KQL ($search)** |
|-------------|-------------------|------------------|
| **Precision** | Exact field matching | Semantic/fuzzy matching |
| **Syntax** | SQL-like operators | Natural language keywords |
| **Performance** | Highly optimized | Good for content search |
| **Use Cases** | Business logic, exact filtering | User search, content discovery |
| **Date Ranges** | ISO format required | Relative dates supported |
| **Boolean Logic** | Full AND/OR/NOT support | Keyword combinations |

**Recommendation**: Use OData for precise business logic, KQL for user-facing search interfaces.

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
Advanced email search with OData filter support and optional field selection.

**Parameters:**
- `user_email` (str): User email
- `query_filter` (str, optional): OData filter
- `top` (int, optional): Maximum number of results (default: 10)
- `folders` (List[str], optional): Folders to search (default: ["Inbox", "SentItems", "Drafts"])
- `fields` (List[str], optional): **NEW** - Specific fields to return (e.g., `['subject', 'sender', 'date', 'id']`)

**🎯 Field Selection Examples:**
```javascript
// Get only essential metadata
fields: ['subject', 'sender', 'date', 'id']

// Get summary information without full body
fields: ['subject', 'sender', 'summary', 'date']

// Get recipients and categorization
fields: ['subject', 'sender', 'cc', 'date', 'id']

// All available fields
fields: ['subject', 'sender', 'date', 'id', 'body', 'summary', 'cc']
```

**OData filter examples (Validated ✅):**
```javascript
// Unread emails
"isRead eq false"

// Emails from specific sender
"from/emailAddress/address eq 'user@example.com'"

// Emails containing keywords in subject
"contains(subject, 'invoice')"

// Emails starting with specific text
"startswith(subject, 'Newsletter')"

// Date range filtering
"receivedDateTime ge 2025-01-01T00:00:00Z and receivedDateTime le 2025-12-31T23:59:59Z"

// High importance unread emails
"isRead eq false and importance eq 'high'"

// Emails with attachments from last month
"hasAttachments eq true and receivedDateTime ge 2025-06-01T00:00:00Z"

// Complex OR conditions
"startswith(subject, 'Invoice') or startswith(subject, 'Receipt')"

// Domain-based filtering
"contains(from/emailAddress/address, 'gmail.com')"
```

**Performance Notes:**
- All OData operators validated: `eq`, `ne`, `gt`, `ge`, `lt`, `le`, `contains`, `startswith`
- Boolean logic fully supported: `and`, `or`, `not`
- Date filtering requires ISO 8601 format with timezone
- Use `_No_Body` variants for 2-3x faster performance when full content isn't needed

#### 2b. Search_Outlook_Emails_No_Body *(Performance Optimized)*
Performance-optimized email search that excludes email body content for faster processing and reduced data transfer.

**Parameters:**
- `user_email` (str): User email
- `query_filter` (str, optional): OData filter (same syntax as Search_Outlook_Emails)
- `top` (int, optional): Maximum number of results (default: 10)
- `folders` (List[str], optional): Folders to search (default: ["Inbox", "SentItems", "Drafts"])
- `fields` (List[str], optional): **NEW** - Specific fields to return (optimized for no-body searches)

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
- `fields` (List[str], optional): **NEW** - Specific fields to return with intelligent KQL/OData mapping

**KQL Search Examples (Validated ✅):**
```javascript
// Search by sender (multiple formats supported)
"from:john@example.com"
"from:John Doe" 

// Content and subject searches
"meeting agenda"
"subject:urgent"

// Recipient-based searches  
"to:maria@company.com"

// Date-based searches (natural language)
"received:today"
"received:last week"
"sent:2025-07-01..2025-07-07"

// Attachment searches
"hasattachment:true"

// Complex keyword combinations
"AI AND newsletter"
"project AND deadline NOT completed"
"invoice AND (Hetzner OR Google)"

// Mixed field and content searches
"from:boss@company.com urgent meeting"
```

**KQL Advantages:**
- Natural language date expressions
- Fuzzy matching and semantic search
- Content-based discovery across subject and body
- User-friendly syntax for end-user interfaces
- Excellent for exploratory searches

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
- `fields` (List[str], optional): **NEW** - Specific fields to return (optimized for KQL no-body searches)

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
      // Note: 'body' field only included in full search functions unless specifically requested in 'fields'
    }
  ]
}
```

**🎯 Field Selection Response Examples:**
```json
// With fields: ['subject', 'sender', 'id']
{
  "data": [
    {
      "asunto": "Project meeting",        // OData results use Spanish keys
      "remitente": "John Doe <john@example.com>",
      "id": "AAMkAGE1M2IyNGNmLWI4..."
    }
  ]
}

// KQL search with fields: ['subject', 'sender', 'date']
{
  "data": [
    {
      "subject": "Project meeting",       // KQL results use English keys
      "from": "John Doe <john@example.com>",
      "receivedDateTime": "2024-01-15 10:30:00"
    }
  ]
}
```

**Field Mapping Reference:**
| Standard Field | OData Key (Spanish) | KQL Key (English) |
|---------------|-------------------|------------------|
| `subject` | `asunto` | `subject` |
| `sender` | `remitente` | `from` |
| `date` | `fecha` | `receivedDateTime` |
| `cc` | `copiados` | `ccRecipients` |
| `body` | `cuerpo` | `body` |
| `summary` | `resumen` | `bodyPreview` |
| `id` | `id` | `id` |

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

## 🆕 Latest Updates (v0.1.15+)

### Major Refactoring & Optimization
- **Unified Resource Architecture**: Complete refactoring of `resources.py` with unified search logic and pagination
- **Eliminated Code Duplication**: Removed legacy functions and consolidated helpers for better maintainability
- **Enhanced Performance**: Optimized data fetching with configurable `include_body` parameter across all search functions
- **Improved Error Handling**: Robust exception handling and fallback strategies
- **Attachment Management**: Streamlined attachment handling with unified API approach

### Advanced Search Capabilities
- **Comprehensive $filter Support**: Full OData operator support validated (eq, ne, gt, ge, lt, le, contains, startswith)
- **Complete KQL Integration**: Natural language search with Microsoft's Keyword Query Language
- **Combined Logic Operations**: Complex AND/OR filtering with multiple conditions
- **Performance Optimized**: Smart body inclusion/exclusion based on use case
- **Extensive Testing**: Thoroughly validated against real mailbox data (sss@sofias.ai)

### 🎯 NEW: Field Selection Feature
- **Selective Data Retrieval**: All search tools now support a `fields` parameter for returning only specific email fields
- **Bandwidth Optimization**: Reduce response size by selecting only needed fields (e.g., `['subject', 'sender', 'date', 'id']`)
- **Dual Mapping Support**: Intelligent field mapping for both OData (Spanish keys) and KQL (English keys) search results
- **Autogen Compatible**: Returns arrays optimized for AI agent frameworks and automated processing
- **Flexible Field Names**: Supports standardized field names (`subject`, `sender`, `date`, `id`, `body`, `summary`, `cc`) across all search methods

### Reply Drafts with Preserved Conversation History
- The `create_draft_email_tool` function supports a `reply_to_id` parameter
- When replying, drafts automatically include full conversation history with original formatting
- HTML and plain text preservation for maximum fidelity
- Seamless integration with Outlook's native reply behavior

#### Example usage
```python
create_draft_email_tool(
    subject="Re: Project Update",
    body="Thank you for the update!",
    to_recipients=["user@example.com"],
    user_email="me@example.com",
    reply_to_id="AAMkAGI2...",
    category="Work"  # Now supports categories in replies
)
```

### Validated Search Examples
All search patterns have been extensively tested and validated:

```python
# Advanced OData filtering
"receivedDateTime ge 2025-07-01T00:00:00Z and isRead eq false"
"contains(subject, 'invoice') and importance eq 'high'"
"startswith(subject, 'Newsletter') or endswith(from/emailAddress/address, 'gmail.com')"

# KQL natural language searches  
"AI newsletter", "from:user@domain.com sent:2025-07-01..2025-07-07"
"subject:urgent AND hasattachment:true"
```

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

### Example 3: Performance-Optimized Searches with Field Selection
```python
# Fast search for email previews without body processing
preview_results = search_emails_no_body_by_search_query_tool(
    user_email="user@company.com",
    search_query="from:boss@company.com project",
    top=50
)

# Get only essential fields for UI display
essential_data = search_emails_tool(
    user_email="user@company.com",
    query_filter="isRead eq false",
    fields=['subject', 'sender', 'date', 'id'],
    top=100
)

# Bandwidth-optimized search for autogen/AI processing
ai_optimized = search_emails_by_search_query_tool(
    user_email="user@company.com",
    search_query="meeting AND urgent",
    fields=['subject', 'sender', 'summary'],
    top=20
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

### Field Selection System (NEW in v0.1.15+)

The server now supports selective field retrieval across all search functions:

1. **Standardized Field Names**: Use consistent field names (`subject`, `sender`, `date`, `id`, `body`, `summary`, `cc`) across all search methods
2. **Intelligent Mapping**: Automatic detection and mapping between OData (Spanish keys) and KQL (English keys) result formats
3. **Bandwidth Optimization**: Return only needed fields to reduce response size and improve performance
4. **Autogen Compatibility**: Optimized array responses for AI agent frameworks and automated processing
5. **Flexible Implementation**: Works with both full and no-body search variants

### Unified Architecture (Latest)

The server features a completely refactored architecture with:

1. **Unified Resource Management**: Single source of truth for all email operations in `resources.py`
2. **Optimized Data Fetching**: Configurable body inclusion with `include_body` parameter
3. **Consolidated Pagination**: Unified pagination logic across all search functions
4. **Eliminated Duplication**: Removed legacy functions and consolidated helper methods
5. **Enhanced Error Handling**: Robust exception handling with detailed logging

### Content Cleaning System

The server includes an advanced cleaning system that:

1. **Converts HTML to plain text** using html2text with optimized configuration
2. **Removes unwanted elements**: scripts, styles, metadata, etc.
3. **Normalizes text**: special characters, spaces, line breaks
4. **Detects and truncates disclaimers**: identifies legal/boilerplate text patterns
5. **Generates automatic summaries**: extracts first 150 relevant characters
6. **Smart Body Processing**: Uses native Graph `bodyPreview` for `_No_Body` functions for optimal performance

### Robust Error Handling

- Exception handling decorators on all critical operations
- Fallback strategies for content processing
- Detailed logging for debugging
- Input parameter validation
- Graceful degradation for malformed content

### Performance Optimizations

- **Smart Body Inclusion**: `include_body=False` for 2-3x performance improvement
- **Field Selection**: NEW - Return only specific fields to reduce bandwidth and processing time
- **Native Graph Features**: Leverages `bodyPreview` for summaries when available
- **Unified Pagination**: Consistent page limits to avoid timeouts
- **Optimized Queries**: Reduced API calls through intelligent batching
- **Memory Efficient**: Processes large result sets without memory bloat
- **Dual Mapping System**: Efficient field mapping for both OData and KQL result formats

## Project Structure

```
mcp-outlook/
├── src/
│   └── mcp_outlook_server/
│       ├── __init__.py
│       ├── server.py             # Server entry point
│       ├── common.py             # Configuration and Graph client setup
│       ├── tools.py              # MCP tool definitions (refactored & optimized)
│       ├── resources.py          # Unified resource access logic (major refactor)
│       ├── format_utils.py       # Email formatting and extraction utilities
│       ├── format_utils_search.py # Search result formatting
│       └── clean_utils.py        # Advanced content cleaning system
├── .env                          # Environment variables (don't include in git)
├── pyproject.toml               # Modern Python project configuration
└── README.md                    # This file
```

### Recent Architecture Changes

- **`resources.py`**: Completely refactored with unified search logic, eliminated duplicate functions, added field selection system
- **`tools.py`**: Streamlined tool definitions, improved attachment handling, consolidated helpers, added `fields` parameter support
- **Performance**: Added configurable `include_body` parameter and field selection across all search functions
- **Field Mapping**: Implemented dual mapping system for OData (Spanish) and KQL (English) result compatibility
- **Testing**: Comprehensive validation documented in conversation history, including field selection scenarios

## Logging

The server generates detailed logs in:
- File: `mcp_outlook.log`
- Console: Standard output

Logging configuration in `common.py` with INFO level by default.

## Troubleshooting

### Authentication Error
- Verify environment variables are correctly configured in `.env` file:
  - `CLIENT_ID`: Your Azure application (client) ID
  - `CLIENT_SECRET`: Your Azure application client secret
  - `TENANT_ID`: Your Azure tenant (directory) ID
- Confirm Azure AD application has necessary permissions
- Check that tenant ID is correct
- Ensure the `.env` file is in the project root directory

### Common Environment Variable Issues
- **Error: "Missing required environment variables"**: Make sure variable names match exactly (`CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`)
- **Error: "Authentication failed"**: Verify the client secret hasn't expired and values are correct

### Search Errors
- Validate OData filter syntax
- Verify user has access to specified folders
- Check logs for specific Graph API errors

### Attachment Issues
- Confirm file paths exist and are accessible
- Verify files don't exceed Outlook size limits (25MB per attachment, 150MB total)
- Check file read permissions
- Ensure attachment functions use the unified API approach (implemented in latest version)

### Performance Issues
- Use `_No_Body` functions for listings and previews (2-3x faster)
- Use `fields` parameter to return only needed data (reduces bandwidth and processing time)
- Limit `top` parameter to reasonable values (10-50 for UI, 100+ for batch processing)
- Consider search scope - searching specific folders is faster than all folders
- For large datasets, implement pagination using multiple calls
- Field selection examples: `['subject', 'sender', 'id']` for lists, `['subject', 'sender', 'summary', 'date']` for previews

### Search Issues  
- **OData syntax**: Ensure proper quoting and escaping of string values
- **Date formats**: Use ISO 8601 format with timezone (e.g., `2025-07-01T00:00:00Z`)
- **Field names**: Use exact field names (`from/emailAddress/address`, not `from.email`)
- **KQL queries**: Don't mix OData and KQL syntax in the same query
- **Field selection**: Use standardized field names (`subject`, `sender`, `date`, `id`, `body`, `summary`, `cc`)
- **Result format**: OData returns Spanish keys, KQL returns English keys - field mapping handles this automatically
- Check the validation results in conversation history for confirmed working examples

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