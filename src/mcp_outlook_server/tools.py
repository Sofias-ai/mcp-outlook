import base64, mimetypes, os
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

@mcp.tool(name="Search_Outlook_Emails", description="Searches emails using OData filter syntax (e.g., 'subject eq \'Update\'', 'isRead eq false', date ranges, sender, etc.).")
@_handle_outlook_operation
def search_emails_tool(user_email: str, query_filter: Optional[str] = None, top: int = 10, folders: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.search_emails(user_email, query_filter, folders, top, structured=True)

@mcp.tool(name="Search_Outlook_Emails_No_Body", description="Busca correos usando filtros OData y devuelve todo menos el cuerpo del correo.")
@_handle_outlook_operation
def search_emails_no_body_tool(user_email: str, query_filter: Optional[str] = None, top: int = 10, folders: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.search_emails_no_body(user_email, query_filter, folders, top, structured=True)

@mcp.tool(name="Search_Outlook_Emails_By_Search_Query", description="Busca correos usando el parámetro $search (KQL), permite búsquedas por destinatario, asunto, etc. Devuelve todos los campos.")
@_handle_outlook_operation
def search_emails_by_search_query_tool(user_email: str, search_query: str, top: int = 10, folders: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.search_emails_by_search_query(user_email, search_query, folders, top, structured=True)

@mcp.tool(name="Search_Outlook_Emails_No_Body_By_Search_Query", description="Busca correos usando el parámetro $search (KQL), permite búsquedas por destinatario, asunto, etc. Devuelve todo menos el cuerpo.")
@_handle_outlook_operation
def search_emails_no_body_by_search_query_tool(user_email: str, search_query: str, top: int = 10, folders: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    return resources.search_emails_no_body_by_search_query(user_email, search_query, folders, top, structured=True)

@mcp.tool(name="Create_Outlook_Draft_Email",description="Creates a new draft email with the specified subject, body, recipients, optional category y adjuntos.")
@_handle_outlook_operation
def create_draft_email_tool(subject: str, body: str, to_recipients: List[str], user_email: str, cc_recipients: Optional[List[str]] = None, bcc_recipients: Optional[List[str]] = None, body_type: str = "HTML", category: Optional[str] = None, file_paths: Optional[List[str]] = None) -> Dict[str, Any]:
    body_content = {"content": body, "contentType": body_type}
    builder = graph_client.users[user_email].messages.add(subject=subject, body=body, to_recipients=to_recipients)
    builder.set_property("body", body_content)
    for attr, value in [("ccRecipients", cc_recipients), ("bccRecipients", bcc_recipients), ("categories", [category] if category else None)]:
        if value: builder.set_property(attr, _fmt(value) if attr.endswith("Recipients") else value)
    if file_paths:
        for path in file_paths:
            raw = open(path, "rb").read()
            b64 = base64.b64encode(raw).decode()
            ctype = mimetypes.guess_type(path)[0] or "application/octet-stream"
            builder = builder.add_file_attachment(os.path.basename(path), content_type=ctype, base64_content=b64)
    draft = builder.execute_query()
    return {"id": draft.id, "web_link": draft.web_link}

@mcp.tool(name="Update_Outlook_Draft_Email",description="Updates an existing draft email specified by su ID, incluyendo adjuntos locales opcionales.")
@_handle_outlook_operation
def update_draft_email_tool(message_id: str, user_email: str, subject: Optional[str] = None, body: Optional[str] = None, to_recipients: Optional[List[str]] = None, cc_recipients: Optional[List[str]] = None, bcc_recipients: Optional[List[str]] = None, body_type: Optional[str] = None, category: Optional[str] = None, file_paths: Optional[List[str]] = None) -> Dict[str, Any]:
    msg = graph_client.users[user_email].messages[message_id].get().execute_query()
    if not getattr(msg, "is_draft", True): raise ValueError("Only draft messages can be updated.")
    for attr, value, transform in [
        ("subject", subject, None),
        ("body", body, lambda b: {"contentType": body_type or "Text", "content": b}),
        ("toRecipients", to_recipients, _fmt),
        ("ccRecipients", cc_recipients, _fmt),
        ("bccRecipients", bcc_recipients, _fmt),
        ("categories", [category] if category else None, None)
    ]:
        if value is not None: msg.set_property(attr, transform(value) if transform else value)
    builder = msg.update()
    if file_paths:
        for path in file_paths:
            raw = open(path, "rb").read()
            b64 = base64.b64encode(raw).decode()
            ctype = mimetypes.guess_type(path)[0] or "application/octet-stream"
            builder = builder.add_file_attachment(os.path.basename(path), content_type=ctype, base64_content=b64)
    updated = builder.execute_query()
    return {"id": updated.id, "web_link": updated.web_link}

@mcp.tool(name="Delete_Outlook_Email", description="Deletes an email by its ID.")
@_handle_outlook_operation
def delete_email_tool(message_id: str, user_email: str) -> Dict[str, Any]:
    graph_client.users[user_email].messages[message_id].delete_object().execute_query()
    return {"message": f"Email {message_id} deleted successfully."}