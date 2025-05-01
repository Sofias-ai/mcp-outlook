import re
from typing import List, Dict, Any, Optional
from datetime import datetime
from .common import logger, graph_client

def _extract_email_address_from_string(text):
    if not text: return None
    email_pattern = r'[\w\.-]+@[\w\.-]+\.\w+'
    match = re.search(email_pattern, text)
    return match.group(0) if match else text

def _format_email_output(message) -> Dict[str, Any]:
    # Extract email from recipient object
    def extract_email_from_recipient(recipient):
        if isinstance(recipient, str):
            return _extract_email_address_from_string(recipient)
        if hasattr(recipient, 'email_address') and hasattr(recipient.email_address, 'address'):
            return recipient.email_address.address
        elif hasattr(recipient, 'address'):
            return recipient.address
        elif hasattr(recipient, 'get_property') and 'address' in dir(recipient):
            return recipient.address
        extracted_email = _extract_email_address_from_string(str(recipient))
        if extracted_email: return extracted_email
        logger.warning(f"Could not extract email from recipient: {recipient}")
        return None
    
    # Extract recipients
    to_recipients = []
    if message.to_recipients:
        for recipient in message.to_recipients:
            to_recipients.append(extract_email_from_recipient(recipient))
    
    cc_recipients = []
    if message.cc_recipients:
        for recipient in message.cc_recipients:
            cc_recipients.append(extract_email_from_recipient(recipient))
    
    bcc_recipients = []
    if message.bcc_recipients:
        for recipient in message.bcc_recipients:
            bcc_recipients.append(extract_email_from_recipient(recipient))
    
    # Extract sender
    sender_email = None
    if hasattr(message, 'sender') and message.sender:
        sender_email = extract_email_from_recipient(message.sender)
    
    if not sender_email and hasattr(message, 'from_'):
        sender_email = extract_email_from_recipient(message.from_)
    
    from_email = None
    if hasattr(message, 'from_property') and message.from_property:
        from_email = extract_email_from_recipient(message.from_property)
    
    # Extract date properties
    def get_datetime_iso(obj, property_names):
        for name in property_names:
            if hasattr(obj, name):
                attr = getattr(obj, name)
                if attr:
                    if hasattr(attr, 'isoformat'): return attr.isoformat()
                    return str(attr)
        return None
    
    received_datetime = get_datetime_iso(message, ['receivedDateTime', 'received_date_time', 'received_datetime', 'dateTimeReceived'])
    sent_datetime = get_datetime_iso(message, ['sentDateTime', 'sent_date_time', 'sent_datetime', 'dateTimeSent'])
    
    # Extract reply-to
    reply_to = []
    if hasattr(message, 'reply_to') and message.reply_to:
        for recipient in message.reply_to:
            email = extract_email_from_recipient(recipient)
            if email: reply_to.append(email)
    
    return {
        "id": message.id,
        "subject": message.subject,
        "body_preview": message.body_preview if hasattr(message, 'body_preview') else None,
        "body": message.body.content if hasattr(message, 'body') and message.body else None,
        "body_type": message.body.contentType if hasattr(message, 'body') and message.body and hasattr(message.body, 'contentType') else None,
        "sender": sender_email or from_email,
        "from": from_email or sender_email,
        "to_recipients": to_recipients,
        "cc_recipients": cc_recipients,
        "bcc_recipients": bcc_recipients,
        "reply_to": reply_to,
        "received_datetime": received_datetime,
        "sent_datetime": sent_datetime,
        "has_attachments": message.has_attachments if hasattr(message, 'has_attachments') else False,
        "importance": message.importance if hasattr(message, 'importance') else None,
        "is_draft": message.is_draft if hasattr(message, 'is_draft') else None,
        "is_read": message.is_read if hasattr(message, 'is_read') else None,
        "conversation_id": message.conversation_id if hasattr(message, 'conversation_id') else None,
        "web_link": message.web_link if hasattr(message, 'web_link') else None,
    }

def get_email_by_id(message_id: str, user_email: str) -> Optional[Dict[str, Any]]:
    logger.info(f"Getting email with ID: {message_id} for user {user_email}")
    message = graph_client.users[user_email].messages[message_id].get().execute_query()
    if message:
        logger.info(f"Successfully retrieved email with ID: {message_id}")
        return _format_email_output(message)
    else:
        logger.warning(f"Email with ID: {message_id} not found.")
        return None

def search_emails(query: str, user_email: str, top: int = 10, folders: List[str] = None) -> List[Dict[str, Any]]:
    if folders is None:
        folders = ["Inbox", "SentItems", "Drafts"]
    
    logger.info(f"Searching emails for user {user_email} with query: '{query}', top: {top}, folders: {folders}")
    
    all_messages = []
    for folder_name in folders:
        folder = graph_client.users[user_email].mail_folders[folder_name].get().execute_query()
        messages = graph_client.users[user_email].mail_folders[folder_name].messages.filter(query).top(top).get().execute_query()
        logger.info(f"Found {len(messages)} emails matching query in {folder_name}.")
        all_messages.extend(messages)
        
        if len(all_messages) >= top:
            all_messages = all_messages[:top]
            break
    
    return [_format_email_output(msg) for msg in all_messages]

def download_emails_by_date(start_date_str: str, end_date_str: str, user_email: str, top: int = 100, folders: List[str] = None) -> List[Dict[str, Any]]:
    if folders is None:
        folders = ["Inbox", "SentItems", "Drafts"]
    
    logger.info(f"Downloading emails for user {user_email} between {start_date_str} and {end_date_str}, top: {top}, folders: {folders}")
    
    all_messages = []
    start_date = datetime.fromisoformat(start_date_str.replace("Z", "+00:00")).strftime('%Y-%m-%dT%H:%M:%SZ')
    end_date = datetime.fromisoformat(end_date_str.replace("Z", "+00:00")).strftime('%Y-%m-%dT%H:%M:%SZ')

    for folder_name in folders:
        date_field = "sentDateTime" if folder_name == "SentItems" else "receivedDateTime"
        query = f"{date_field} ge {start_date} and {date_field} le {end_date}"
        logger.info(f"Using query for {folder_name}: {query}")
        
        messages = graph_client.users[user_email].mail_folders[folder_name].messages.filter(query).top(top).get().execute_query()
        logger.info(f"Found {len(messages)} emails in date range in {folder_name}.")
        all_messages.extend(messages)
        
        if len(all_messages) >= top:
            all_messages = all_messages[:top]
            break
    
    return [_format_email_output(msg) for msg in all_messages]