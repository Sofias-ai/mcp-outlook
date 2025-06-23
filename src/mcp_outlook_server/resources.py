from typing import List, Any, Optional
from .common import graph_client, _get_graph_access_token
from .format_utils import format_email_output
from .clean_utils import format_email_structured, format_emails_list_structured, format_email_structured_no_body, format_emails_list_structured_no_body
from .format_utils_search import format_emails_list_structured_search, format_emails_list_structured_search_no_body
import json
import requests
from office365.runtime.http.request_options import RequestOptions

def _fetch_emails(user_email: str, query_filter: Optional[str], folders: Optional[List[str]], top: int, select_fields: Optional[List[str]] = None) -> List[Any]:
    # Internal helper to fetch emails with optional $select fields
    if not folders:
        folders = ["Inbox", "SentItems", "Drafts"]
    all_messages = []
    for folder_name in folders:
        if len(all_messages) >= top:
            break
        query_obj = graph_client.users[user_email].mail_folders[folder_name].messages
        if select_fields:
            query_obj = query_obj.select(select_fields)
        if query_filter:
            query_obj = query_obj.filter(query_filter)
        page_collection = query_obj.paged().top(1000).get().execute_query()
        messages = list(page_collection)
        remaining_space = top - len(all_messages)
        all_messages.extend(messages[:remaining_space] if remaining_space < len(messages) else messages)
        page_count = 1
        while page_collection.has_next and len(all_messages) < top and page_count <= 20:
            page_collection = page_collection.get().execute_query()
            messages = list(page_collection)
            if not messages:
                break
            remaining_needed = top - len(all_messages)
            all_messages.extend(messages[:remaining_needed] if remaining_needed < len(messages) else messages)
            page_count += 1
            if len(all_messages) >= top:
                break
        if len(all_messages) >= top:
            break
    return all_messages

def _fetch_emails_by_search(user_email: str, search_query: str, folders: Optional[List[str]], top: int, select_fields: Optional[List[str]] = None) -> List[Any]:
    # Internal helper to fetch emails using $search (KQL), with optional $select fields, via direct HTTP
    if not folders:
        folders = ["Inbox", "SentItems", "Drafts"]
    all_messages = []
    access_token = _get_graph_access_token()
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    for folder_name in folders:
        if len(all_messages) >= top:
            break
        url = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/{folder_name}/messages"
        params = {
            "$search": f'"{search_query}"',
            "$top": str(min(1000, top))
        }
        if select_fields:
            params["$select"] = ",".join(select_fields)
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        messages = data.get("value", [])
        remaining_space = top - len(all_messages)
        all_messages.extend(messages[:remaining_space] if remaining_space < len(messages) else messages)
        if len(all_messages) >= top:
            break
    return all_messages

def search_emails(user_email: str, query_filter: Optional[str] = None, folders: Optional[List[str]] = None, top: int = 10, as_text: bool = True, structured: bool = True) -> List[Any]:
    # Standard email search, fetches all fields (including body) and applies cleaning/formatting
    all_messages = _fetch_emails(user_email, query_filter, folders, top, select_fields=None)  # No $select, trae todo
    return format_emails_list_structured(all_messages) if structured else ([format_email_output(msg, as_text=True) for msg in all_messages] if as_text else [format_email_output(msg, as_text=False) for msg in all_messages])

def get_email_by_id(message_id: str, user_email: str, as_text: bool = True, structured: bool = True) -> Optional[Any]:
    # Get a specific email by its ID
    message = graph_client.users[user_email].messages[message_id].get().execute_query()
    return format_email_structured(message) if structured else format_email_output(message, as_text=as_text)

def search_emails_no_body(user_email: str, query_filter: Optional[str] = None, folders: Optional[List[str]] = None, top: int = 10, as_text: bool = True, structured: bool = True) -> List[Any]:
    # Same as search_emails but optimized: only fetches required fields (without body), uses bodyPreview for summary.
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
    return format_emails_list_structured_no_body(all_messages) if structured else ([format_email_output(msg, as_text=True) for msg in all_messages] if as_text else [format_email_output(msg, as_text=False) for msg in all_messages])


def search_emails_by_search_query(user_email: str, search_query: str, folders: Optional[List[str]] = None, top: int = 10, as_text: bool = True, structured: bool = True) -> List[Any]:
    # Search emails using $search (KQL), fetches all fields and applies cleaning/formatting
    all_messages = _fetch_emails_by_search(user_email, search_query, folders, top)
    return format_emails_list_structured_search(all_messages) if structured else (
        [format_email_output(msg, as_text=True) for msg in all_messages] if as_text else [format_email_output(msg, as_text=False) for msg in all_messages]
    )

def search_emails_no_body_by_search_query(user_email: str, search_query: str, folders: Optional[List[str]] = None, top: int = 10, as_text: bool = True, structured: bool = True) -> List[Any]:
    # Search emails using $search (KQL), only fetches required fields (without body), uses bodyPreview for summary.
    select_fields = [
        "id", "subject", "sender", "from", "toRecipients", "ccRecipients", "bccRecipients",
        "receivedDateTime", "sentDateTime", "isRead", "importance", "categories", "bodyPreview"
    ]
    all_messages = _fetch_emails_by_search(user_email, search_query, folders, top, select_fields=select_fields)
    return format_emails_list_structured_search_no_body(all_messages) if structured else (
        [format_email_output(msg, as_text=True) for msg in all_messages] if as_text else [format_email_output(msg, as_text=False) for msg in all_messages]
    )

def get_raw_email_by_id(message_id: str, user_email: str) -> Any:
    # Devuelve el email tal cual, sin limpiar ni formatear
    return graph_client.users[user_email].messages[message_id].get().execute_query()