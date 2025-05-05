from typing import List, Any, Optional, Dict
from datetime import datetime
from .common import logger, graph_client
from .format_utils import format_email_output

def _query_emails(user_email: str, query_filter: str = None, folder_names: List[str] = None, 
                top: int = 10, as_text: bool = True) -> List[Any]:
    
    """Base function for querying emails with various filters"""
    if not folder_names:folder_names = ["Inbox", "SentItems", "Drafts"]
    logger.info(f"Querying emails for {user_email} with filter: {query_filter}, folders: {folder_names}, top: {top}")
    
    all_messages = []
    for folder_name in folder_names:
        
        query_obj = graph_client.users[user_email].mail_folders[folder_name].messages
        if query_filter: query_obj = query_obj.filter(query_filter)
        
        messages = query_obj.top(top).get().execute_query()
        logger.info(f"Found {len(messages)} emails in {folder_name}")
        all_messages.extend(messages)
        
        if len(all_messages) >= top:
            all_messages = all_messages[:top]
            break
    
    # Apply formatting as needed
    return [format_email_output(msg, as_text=as_text) for msg in all_messages]

def get_email_by_id(message_id: str, user_email: str, as_text: bool = True) -> Optional[Any]:
    """Gets a specific email by its ID"""
    logger.info(f"Getting email with ID: {message_id} for user {user_email}")
    message = graph_client.users[user_email].messages[message_id].get().execute_query()
    
    if message:
        logger.info(f"Successfully retrieved email with ID: {message_id}")
        return format_email_output(message, as_text=as_text)
    
    logger.warning(f"Email with ID: {message_id} not found")
    return None

def search_emails(query: str, user_email: str, top: int = 10, folders: List[str] = None, as_text: bool = True) -> List[Any]:
    """Searches emails using OData queries"""
    return _query_emails(user_email, query, folders, top, as_text)

def download_emails_by_date(start_date_str: str, end_date_str: str, user_email: str, 
                          top: int = 100, folders: List[str] = None, as_text: bool = True) -> List[Any]:
    """Downloads emails within a date range"""
    if not folders:
        folders = ["Inbox", "SentItems", "Drafts"]
    
    start_date = datetime.fromisoformat(start_date_str.replace("Z", "+00:00")).strftime('%Y-%m-%dT%H:%M:%SZ')
    end_date = datetime.fromisoformat(end_date_str.replace("Z", "+00:00")).strftime('%Y-%m-%dT%H:%M:%SZ')
    
    all_messages = []
    for folder_name in folders:
        date_field = "sentDateTime" if folder_name == "SentItems" else "receivedDateTime"
        query = f"{date_field} ge {start_date} and {date_field} le {end_date}"
        
        messages = _query_emails(user_email, query, [folder_name], top, False)
        all_messages.extend(messages)
        
        if len(all_messages) >= top:
            all_messages = all_messages[:top]
            break
    
    # Add summary and apply formatting
    results = []
    if all_messages:
        summary = f"Found {len(all_messages)} emails in the inbox from {start_date_str} to {end_date_str}."
        results.append(summary)
        
        if as_text:
            results.extend([format_email_output(msg, as_text=True) for msg in all_messages])
        else:
            results = all_messages
    
    return results