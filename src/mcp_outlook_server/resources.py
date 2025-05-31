from typing import List, Any, Optional
from datetime import datetime
from .common import logger, graph_client
from .format_utils import format_email_output
from .clean_utils import format_email_structured, format_emails_list_structured

def _query_emails(user_email: str, query_filter: Optional[str] = None, folder_names: Optional[List[str]] = None, top: int = 10, as_text: bool = True, structured: bool = True) -> List[Any]:
    # Base function for querying emails with various filters
    if not folder_names: folder_names = ["Inbox", "SentItems", "Drafts"]
    logger.info(f"Querying emails for {user_email} with filter: {query_filter}, folders: {folder_names}, top: {top}")
    
    all_messages = []
    for folder_name in folder_names:
        query_obj = graph_client.users[user_email].mail_folders[folder_name].messages
        if query_filter: query_obj = query_obj.filter(query_filter)
        messages = query_obj.top(top).get().execute_query()
        logger.info(f"Found {len(messages)} emails in {folder_name}")
        all_messages.extend(messages)
        if len(all_messages) >= top: all_messages = all_messages[:top]; break
    
    if structured: return format_emails_list_structured(all_messages)
    elif as_text: return [format_email_output(msg, as_text=True) for msg in all_messages]
    else: return [format_email_output(msg, as_text=False) for msg in all_messages]

def get_email_by_id(message_id: str, user_email: str, as_text: bool = True, structured: bool = True) -> Optional[Any]:
    # Get a specific email by its ID
    logger.info(f"Getting email with ID: {message_id} for user {user_email}")
    message = graph_client.users[user_email].messages[message_id].get().execute_query()
    if message:
        logger.info(f"Successfully retrieved email with ID: {message_id}")
        return format_email_structured(message) if structured else format_email_output(message, as_text=as_text)
    logger.warning(f"Email with ID: {message_id} not found")
    return None

def search_emails(query: str, user_email: str, top: int = 10, folders: Optional[List[str]] = None, as_text: bool = True, structured: bool = True) -> List[Any]:
    # Search emails using OData queries
    return _query_emails(user_email, query, folders, top, as_text, structured)

def download_emails_by_date(start_date_str: str, end_date_str: str, user_email: str, top: int = 100, folders: Optional[List[str]] = None, as_text: bool = True, structured: bool = True) -> List[Any]:
    # Download emails within a date range
    if not folders: folders = ["Inbox", "SentItems", "Drafts"]
    start_date = datetime.fromisoformat(start_date_str.replace("Z", "+00:00")).strftime('%Y-%m-%dT%H:%M:%SZ')
    end_date = datetime.fromisoformat(end_date_str.replace("Z", "+00:00")).strftime('%Y-%m-%dT%H:%M:%SZ')
    
    all_messages = []
    for folder_name in folders:
        date_field = "sentDateTime" if folder_name == "SentItems" else "receivedDateTime"
        query = f"{date_field} ge {start_date} and {date_field} le {end_date}"
        messages = _query_emails(user_email, query, [folder_name], top, as_text, structured)
        all_messages.extend(messages)
        if len(all_messages) >= top: all_messages = all_messages[:top]; break
    
    if structured: return all_messages
    else:
        results = []
        if all_messages:
            summary = f"Found {len(all_messages)} emails in the inbox from {start_date_str} to {end_date_str}."
            results.append(summary); results.extend(all_messages)
        return results