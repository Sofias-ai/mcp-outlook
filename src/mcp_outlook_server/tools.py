from functools import wraps
from typing import Optional, List, Dict, Any
from .common import mcp, graph_client, _fmt
from . import resources

def _handle_outlook_operation(func):
    @wraps(func)
    async def wrapper(*args, **kwargs):
        result = func(*args, **kwargs)
        if isinstance(result, (list, type(None))): return {"success": True, "data": result}
        elif isinstance(result, dict) and "success" not in result and func.__name__.startswith(('create', 'update', 'delete')): return {**result, "success": True}
        return result
    return wrapper

@mcp.tool(name="Get_Outlook_Email", description="Retrieves a specific email by its unique ID.")
@_handle_outlook_operation
def get_email_tool(message_id: str, user_email: str) -> Optional[Dict[str, Any]]:
    return resources.get_email_by_id(message_id, user_email, structured=True)

@mcp.tool(name="Search_Outlook_Emails", description="Searches emails using OData filter syntax (e.g., \"subject eq 'Update'\", \"isRead eq false\").")
@_handle_outlook_operation
def search_emails_tool(query: str, user_email: str, top: int = 10, folders: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.search_emails(query, user_email, top, folders or [], structured=True)

@mcp.tool(name="Download_Outlook_Emails_By_Date", description="Downloads emails received within a specific date range (ISO 8601 format: YYYY-MM-DDTHH:MM:SSZ).")
@_handle_outlook_operation
def download_emails_by_date_tool(start_date: str, end_date: str, user_email: str, top: int = 100, folders: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.download_emails_by_date(start_date, end_date, user_email, top, folders or [], structured=True)

@mcp.tool(name="Create_Outlook_Draft_Email", description="Creates a new draft email with the specified subject, body, and recipients.")
@_handle_outlook_operation
def create_draft_email_tool(subject: str, body: str, to_recipients: List[str], user_email: str, cc_recipients: Optional[List[str]] = None, bcc_recipients: Optional[List[str]] = None, body_type: str = "HTML") -> Dict[str, Any]:
    draft = graph_client.users[user_email].messages.add(subject=subject, body=body, to_recipients=to_recipients).execute_query()
    for attr, value in [("ccRecipients", cc_recipients), ("bccRecipients", bcc_recipients)]:
        if value: draft.set_property(attr, _fmt(value))
    return {"id": draft.id, "web_link": draft.web_link}

@mcp.tool(name="Update_Outlook_Draft_Email", description="Updates an existing draft email specified by its ID.")
@_handle_outlook_operation
def update_draft_email_tool(message_id: str, user_email: str, subject: Optional[str] = None, body: Optional[str] = None, to_recipients: Optional[List[str]] = None, cc_recipients: Optional[List[str]] = None, bcc_recipients: Optional[List[str]] = None, body_type: Optional[str] = None) -> Dict[str, Any]:
    msg = graph_client.users[user_email].messages[message_id].get().execute_query()
    if not getattr(msg, "is_draft", True): raise ValueError("Only draft messages can be updated.")
    for attr, value, transform in [("subject", subject, None), ("body", body, lambda b: {"contentType": body_type or "Text", "content": b}), ("toRecipients", to_recipients, _fmt), ("ccRecipients", cc_recipients, _fmt), ("bccRecipients", bcc_recipients, _fmt)]:
        if value is not None: msg.set_property(attr, transform(value) if transform else value)
    updated = msg.update().execute_query()
    return {"id": updated.id, "web_link": updated.web_link}

@mcp.tool(name="Delete_Outlook_Email", description="Deletes an email by its ID.")
@_handle_outlook_operation
def delete_email_tool(message_id: str, user_email: str) -> Dict[str, Any]:
    graph_client.users[user_email].messages[message_id].delete_object().execute_query()
    return {"message": f"Email {message_id} deleted successfully."}