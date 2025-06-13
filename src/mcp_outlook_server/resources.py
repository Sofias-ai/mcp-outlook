from typing import List, Any, Optional
from .common import graph_client
from .format_utils import format_email_output
from .clean_utils import format_email_structured, format_emails_list_structured

def search_emails(user_email: str, query_filter: Optional[str] = None, folders: Optional[List[str]] = None, top: int = 10, as_text: bool = True, structured: bool = True) -> List[Any]:
    if not folders: folders = ["Inbox", "SentItems", "Drafts"]
    all_messages = []
    for folder_name in folders:
        if len(all_messages) >= top: break
        try:
            query_obj = graph_client.users[user_email].mail_folders[folder_name].messages
            if query_filter: query_obj = query_obj.filter(query_filter)
            page_collection = query_obj.paged().top(1000).get().execute_query()
            messages = list(page_collection)
            remaining_space = top - len(all_messages)
            all_messages.extend(messages[:remaining_space] if remaining_space < len(messages) else messages)
            page_count = 1
            while page_collection.has_next and len(all_messages) < top and page_count <= 20:
                try:
                    page_collection = page_collection.get().execute_query()
                    messages = list(page_collection)
                    if not messages: break
                    remaining_needed = top - len(all_messages)
                    all_messages.extend(messages[:remaining_needed] if remaining_needed < len(messages) else messages)
                    page_count += 1
                    if len(all_messages) >= top: break
                except Exception: break
            if len(all_messages) >= top: break
        except Exception: continue
    return format_emails_list_structured(all_messages) if structured else ([format_email_output(msg, as_text=True) for msg in all_messages] if as_text else [format_email_output(msg, as_text=False) for msg in all_messages])

def get_email_by_id(message_id: str, user_email: str, as_text: bool = True, structured: bool = True) -> Optional[Any]:
    """Get a specific email by its ID."""
    try:
        message = graph_client.users[user_email].messages[message_id].get().execute_query()
        return format_email_structured(message) if structured else format_email_output(message, as_text=as_text)
    except Exception:
        return None

def search_emails_no_body(user_email: str, query_filter: Optional[str] = None, folders: Optional[List[str]] = None, top: int = 10, as_text: bool = True, structured: bool = True) -> List[Any]:
    """Same as search_emails but removes the 'body' and 'cuerpo' fields from each email in the result."""
    emails = search_emails(user_email, query_filter, folders, top, as_text, structured)
    if isinstance(emails, list):
        for email in emails:
            if isinstance(email, dict):
                email.pop('body', None)
                email.pop('cuerpo', None)
    return emails