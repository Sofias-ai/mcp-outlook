from typing import List, Any, Optional, Callable
from .common import graph_client, _get_graph_access_token
from .format_utils import format_email_output
from .clean_utils import format_email_structured, format_emails_list_structured, format_emails_list_structured_no_body
from .format_utils_search import format_emails_list_structured_search, format_emails_list_structured_search_no_body
import requests

DEFAULT_FOLDERS = ["Inbox", "SentItems", "Drafts"]
SELECT_FIELDS_NO_BODY = ["id", "subject", "sender", "from", "toRecipients", "ccRecipients", "bccRecipients", 
                        "receivedDateTime", "sentDateTime", "isRead", "importance", "categories", "bodyPreview"]

def _paginate_messages(initial_collection, top: int, max_pages: int = 20) -> List[Any]:
    """Unified pagination logic for message collections"""
    messages, page_count = list(initial_collection), 1
    while initial_collection.has_next and len(messages) < top and page_count <= max_pages:
        try:
            initial_collection = initial_collection.get().execute_query()
            new_messages = list(initial_collection)
            if not new_messages: break
            messages.extend(new_messages[:top - len(messages)])
            page_count += 1
        except Exception: break
    return messages[:top]

def _fetch_from_folders(user_email: str, folders: List[str], top: int, fetch_func: Callable) -> List[Any]:
    """Unified folder iteration with early termination"""
    all_messages = []
    for folder_name in folders:
        if len(all_messages) >= top: break
        try:
            messages = fetch_func(user_email, folder_name, top - len(all_messages))
            all_messages.extend(messages)
        except Exception: continue
    return all_messages[:top]

def _fetch_emails(user_email: str, query_filter: Optional[str], folders: Optional[List[str]], top: int, select_fields: Optional[List[str]] = None) -> List[Any]:
    def fetch_from_folder(user_email: str, folder_name: str, remaining: int) -> List[Any]:
        query_obj = graph_client.users[user_email].mail_folders[folder_name].messages
        if select_fields: query_obj = query_obj.select(select_fields)
        if query_filter: query_obj = query_obj.filter(query_filter)
        collection = query_obj.paged().top(min(1000, remaining)).get().execute_query()
        return _paginate_messages(collection, remaining)
    
    return _fetch_from_folders(user_email, folders or DEFAULT_FOLDERS, top, fetch_from_folder)

def _fetch_emails_by_search(user_email: str, search_query: str, folders: Optional[List[str]], top: int, select_fields: Optional[List[str]] = None) -> List[Any]:
    def fetch_from_folder(user_email: str, folder_name: str, remaining: int) -> List[Any]:
        headers = {"Authorization": f"Bearer {_get_graph_access_token()}", "Content-Type": "application/json"}
        params = {"$search": f'"{search_query}"', "$top": str(min(1000, remaining))}
        if select_fields: params["$select"] = ",".join(select_fields)
        
        resp = requests.get(f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/{folder_name}/messages", 
                           headers=headers, params=params)
        resp.raise_for_status()
        return resp.json().get("value", [])[:remaining]
    
    return _fetch_from_folders(user_email, folders or DEFAULT_FOLDERS, top, fetch_from_folder)

def _format_results(messages: List[Any], structured: bool, as_text: bool, formatter_structured, formatter_unstructured=None) -> List[Any]:
    """Unified result formatting logic"""
    if structured: return formatter_structured(messages)
    if formatter_unstructured: return formatter_unstructured(messages)
    return [format_email_output(msg, as_text=as_text) for msg in messages]

def search_emails(user_email: str, query_filter: Optional[str] = None, folders: Optional[List[str]] = None, 
                 top: int = 10, as_text: bool = True, structured: bool = True, include_body: bool = True) -> List[Any]:
    """Unified email search with OData filters"""
    select_fields = None if include_body else SELECT_FIELDS_NO_BODY
    messages = _fetch_emails(user_email, query_filter, folders, top, select_fields)
    formatter = format_emails_list_structured if include_body else format_emails_list_structured_no_body
    return _format_results(messages, structured, as_text, formatter)

def search_emails_by_search_query(user_email: str, search_query: str, folders: Optional[List[str]] = None, 
                                 top: int = 10, as_text: bool = True, structured: bool = True, include_body: bool = True) -> List[Any]:
    """Unified email search with KQL queries"""
    select_fields = None if include_body else SELECT_FIELDS_NO_BODY
    messages = _fetch_emails_by_search(user_email, search_query, folders, top, select_fields)
    formatter = format_emails_list_structured_search if include_body else format_emails_list_structured_search_no_body
    return _format_results(messages, structured, as_text, formatter)

def get_email_by_id(message_id: str, user_email: str, as_text: bool = True, structured: bool = True) -> Optional[Any]:
    message = graph_client.users[user_email].messages[message_id].get().execute_query()
    return format_email_structured(message) if structured else format_email_output(message, as_text=as_text)