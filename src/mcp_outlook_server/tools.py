from functools import wraps
from typing import Optional, List, Dict, Any
from .common import mcp, graph_client, USER_EMAIL, _fmt
from . import resources

# Error handling decorator
def _handle_outlook_operation(func):
    @wraps(func)
    async def wrapper(*args, **kwargs):
        result = func(*args, **kwargs)
        if isinstance(result, dict) and "success" not in result and func.__name__.startswith(('create', 'update', 'delete')):
            result["success"] = True
        elif isinstance(result, list) or result is None:
            return {"success": True, "data": result}
        return result
    return wrapper

# Reading tools
@mcp.tool(name="Get_Outlook_Email", description="Retrieves a specific email by its unique ID.")
@_handle_outlook_operation
def get_email_tool(message_id: str) -> Optional[Dict[str, Any]]:
    return resources.get_email_by_id(message_id)

@mcp.tool(name="Search_Outlook_Emails", description="Searches emails using OData filter syntax (e.g., \"subject eq 'Update'\", \"isRead eq false\"). Returns a list of emails from Inbox, Sent Items, and Drafts folders.")
@_handle_outlook_operation
def search_emails_tool(query: str, top: int = 10, folders: List[str] = None) -> List[Dict[str, Any]]:
    return resources.search_emails(query, top, folders)

@mcp.tool(name="Download_Outlook_Emails_By_Date", description="Downloads emails received within a specific date range (ISO 8601 format: YYYY-MM-DDTHH:MM:SSZ). Returns a list of emails from Inbox, Sent Items, and Drafts folders.")
@_handle_outlook_operation
def download_emails_by_date_tool(start_date: str, end_date: str, top: int = 100, folders: List[str] = None) -> List[Dict[str, Any]]:
    return resources.download_emails_by_date(start_date, end_date, top, folders)

# Writing/Modification tools
@mcp.tool(name="Create_Outlook_Draft_Email", description="Creates a new draft email with the specified subject, body, and recipients.")
@_handle_outlook_operation
def create_draft_email_tool(subject: str, body: str, to_recipients: List[str],
                            cc_recipients: Optional[List[str]] = None,
                            bcc_recipients: Optional[List[str]] = None,
                            body_type: str = "HTML") -> Dict[str, Any]:
    draft = (graph_client.users[USER_EMAIL].messages.add(
        subject=subject, body=body, to_recipients=to_recipients
    ).execute_query())
    
    if cc_recipients is not None:
        draft.set_property("ccRecipients", _fmt(cc_recipients))
    if bcc_recipients is not None:
        draft.set_property("bccRecipients", _fmt(bcc_recipients))

    return {"id": draft.id, "web_link": draft.web_link}

@mcp.tool(name="Update_Outlook_Draft_Email", description="Updates an existing draft email specified by its ID. Only provided fields are updated.")
@_handle_outlook_operation
def update_draft_email_tool(message_id: str, subject: Optional[str] = None, body: Optional[str] = None,
                            to_recipients: Optional[List[str]] = None,
                            cc_recipients: Optional[List[str]] = None,
                            bcc_recipients: Optional[List[str]] = None,
                            body_type: Optional[str] = None) -> Dict[str, Any]:
    # Get the draft message
    msg = graph_client.users[USER_EMAIL].messages[message_id].get().execute_query()
    
    # Verify it's a draft
    if not getattr(msg, "is_draft", True):
        raise ValueError("Only draft messages can be updated.")

    # Apply changes via set_property
    if subject is not None:
        msg.set_property("subject", subject)
    if body is not None:
        msg.set_property("body", {"contentType": body_type or "Text", "content": body})
    if to_recipients is not None:
        msg.set_property("toRecipients", _fmt(to_recipients))
    if cc_recipients is not None:
        msg.set_property("ccRecipients", _fmt(cc_recipients))
    if bcc_recipients is not None:
        msg.set_property("bccRecipients", _fmt(bcc_recipients))

    # Execute PATCH in Graph
    updated = msg.update().execute_query()
    return {"id": updated.id, "web_link": updated.web_link}

@mcp.tool(name="Delete_Outlook_Email", description="Deletes an email by its ID (moves it to the Deleted Items folder).")
@_handle_outlook_operation
def delete_email_tool(message_id: str) -> Dict[str, Any]:
    message = graph_client.users[USER_EMAIL].messages[message_id]
    message.delete_object().execute_query()
    return {"success": True, "message": f"Email {message_id} deleted successfully."}