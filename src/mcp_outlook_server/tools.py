import base64, mimetypes, os, requests
from functools import wraps
from typing import Optional, List, Dict, Any
from .common import mcp, graph_client, _fmt, _get_graph_access_token
from . import resources

def _handle_outlook_operation(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        result = func(*args, **kwargs)
        if isinstance(result, list): return result
        elif result is None: return []
        elif isinstance(result, dict): return result
        return result
    return wrapper

def _add_attachments_via_api(user_email: str, message_id: str, file_paths: List[str]):
    headers = {"Authorization": f"Bearer {_get_graph_access_token()}", "Content-Type": "application/json"}
    for path in file_paths:
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        requests.post(f"https://graph.microsoft.com/v1.0/users/{user_email}/messages/{message_id}/attachments",
                     headers=headers, json={"@odata.type": "#microsoft.graph.fileAttachment", "name": os.path.basename(path),
                     "contentType": mimetypes.guess_type(path)[0] or "application/octet-stream", "contentBytes": b64}).raise_for_status()

def _set_message_properties(msg_or_builder, cc_recipients, bcc_recipients, category):
    for attr, value in [("ccRecipients", cc_recipients), ("bccRecipients", bcc_recipients), ("categories", [category] if category else None)]:
        if value: msg_or_builder.set_property(attr, _fmt(value) if attr.endswith("Recipients") else value)

def _add_attachments_via_builder(builder, file_paths: List[str]):
    for path in file_paths:
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        builder = builder.add_file_attachment(os.path.basename(path), content_type=mimetypes.guess_type(path)[0] or "application/octet-stream", base64_content=b64)
    return builder

@mcp.tool(name="Get_Outlook_Email", description="Retrieves a specific email by its unique ID.")
@_handle_outlook_operation
def get_email_tool(message_id: str, user_email: str) -> Optional[Dict[str, Any]]:
    return resources.get_email_by_id(message_id, user_email, structured=True)

@mcp.tool(name="Search_Outlook_Emails", description="Searches emails using OData filter syntax. Optional 'fields' param to select specific fields: ['subject', 'sender', 'date', 'id', 'body', 'summary'].")
@_handle_outlook_operation
def search_emails_tool(user_email: str, query_filter: Optional[str] = None, top: int = 10, folders: Optional[List[str]] = None, fields: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.search_emails(user_email, query_filter, folders, top, structured=True, fields=fields)

@mcp.tool(name="Search_Outlook_Emails_No_Body", description="Searches emails using OData filters without body content. Optional 'fields' param: ['subject', 'sender', 'date', 'id', 'summary'].")
@_handle_outlook_operation
def search_emails_no_body_tool(user_email: str, query_filter: Optional[str] = None, top: int = 10, folders: Optional[List[str]] = None, fields: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.search_emails(user_email, query_filter, folders, top, structured=True, include_body=False, fields=fields)

@mcp.tool(name="Search_Outlook_Emails_By_Search_Query", description="Searches emails using KQL (Keyword Query Language). Optional 'fields' param: ['subject', 'sender', 'date', 'id', 'body', 'summary'].")
@_handle_outlook_operation
def search_emails_by_search_query_tool(user_email: str, search_query: str, top: int = 10, folders: Optional[List[str]] = None, fields: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.search_emails_by_search_query(user_email, search_query, folders, top, structured=True, fields=fields)

@mcp.tool(name="Search_Outlook_Emails_No_Body_By_Search_Query", description="Searches emails using KQL without body content. Optional 'fields' param: ['subject', 'sender', 'date', 'id', 'summary'].")
@_handle_outlook_operation
def search_emails_no_body_by_search_query_tool(user_email: str, search_query: str, top: int = 10, folders: Optional[List[str]] = None, fields: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.search_emails_by_search_query(user_email, search_query, folders, top, structured=True, include_body=False, fields=fields)

@mcp.tool(name="Create_Outlook_Draft_Email",description="Creates a new draft email, optionally replying to a message and including its full content as history.")
@_handle_outlook_operation
def create_draft_email_tool(subject: str, body: str, to_recipients: List[str], user_email: str, cc_recipients: Optional[List[str]] = None, bcc_recipients: Optional[List[str]] = None, body_type: str = "HTML", category: Optional[str] = None, file_paths: Optional[List[str]] = None, reply_to_id: Optional[str] = None) -> Dict[str, Any]:
    if reply_to_id:
        headers = {"Authorization": f"Bearer {_get_graph_access_token()}", "Content-Type": "application/json"}
        draft_id = requests.post(f"https://graph.microsoft.com/v1.0/users/{user_email}/messages/{reply_to_id}/createReply", 
                               headers=headers, json={"comment": body}).json()["id"]
        
        if any([cc_recipients, bcc_recipients, category]):
            msg = graph_client.users[user_email].messages[draft_id].get().execute_query()
            _set_message_properties(msg, cc_recipients, bcc_recipients, category)
            msg.update().execute_query()
        
        if file_paths: _add_attachments_via_api(user_email, draft_id, file_paths)
        return {"id": draft_id, "web_link": f"https://outlook.office365.com/owa/?ItemID={draft_id}&exvsurl=1&viewmodel=ReadMessageItem"}
    
    builder = graph_client.users[user_email].messages.add(subject=subject, to_recipients=to_recipients)
    builder.set_property("body", {"content": body, "contentType": body_type})
    _set_message_properties(builder, cc_recipients, bcc_recipients, category)
    if file_paths: builder = _add_attachments_via_builder(builder, file_paths)
    draft = builder.execute_query()
    return {"id": draft.id, "web_link": getattr(draft, 'web_link', None)}

@mcp.tool(name="Update_Outlook_Draft_Email",description="Updates an existing draft email by its ID, including optional local attachments.")
@_handle_outlook_operation
def update_draft_email_tool(message_id: str, user_email: str, subject: Optional[str] = None, body: Optional[str] = None, to_recipients: Optional[List[str]] = None, cc_recipients: Optional[List[str]] = None, bcc_recipients: Optional[List[str]] = None, body_type: Optional[str] = None, category: Optional[str] = None, file_paths: Optional[List[str]] = None) -> Dict[str, Any]:
    msg = graph_client.users[user_email].messages[message_id].get().execute_query()
    if not getattr(msg, "is_draft", True): raise ValueError("Only draft messages can be updated.")
    
    for attr, value, transform in [("subject", subject, None), ("body", body, lambda b: {"contentType": body_type or "Text", "content": b}), 
                                  ("toRecipients", to_recipients, _fmt), ("ccRecipients", cc_recipients, _fmt), 
                                  ("bccRecipients", bcc_recipients, _fmt), ("categories", [category] if category else None, None)]:
        if value is not None: msg.set_property(attr, transform(value) if transform else value)
    
    updated = msg.update().execute_query()
    if file_paths: _add_attachments_via_api(user_email, message_id, file_paths)
    return {"id": updated.id, "web_link": getattr(updated, 'web_link', None)}

@mcp.tool(name="Delete_Outlook_Email", description="Deletes an email by its ID.")
@_handle_outlook_operation
def delete_email_tool(message_id: str, user_email: str) -> Dict[str, Any]:
    graph_client.users[user_email].messages[message_id].delete_object().execute_query()
    return {"message": f"Email {message_id} deleted successfully."}