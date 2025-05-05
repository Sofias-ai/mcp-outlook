from typing import Any, Tuple
from .clean_utils import clean_email_content
import re, json, logging
from datetime import datetime
from functools import wraps

# Configure logging
logger = logging.getLogger(__name__)

def safe_operation(default_return=None):
    """Decorator to handle exceptions in attribute access operations"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                logger.debug(f"Error in {func.__name__}: {e}")
                return default_return
        return wrapper
    return decorator

@safe_operation()
def get_message_attribute(obj: Any, paths: list, default=None) -> Any:
    """Unified function to get attributes from message objects"""
    if not obj:
        return default
    
    # Fast path for direct attribute access
    for path in paths:
        # Handle nested attributes
        if '.' in path:
            parts = path.split('.')
            current = obj
            for part in parts:
                if not hasattr(current, part) or (current := getattr(current, part)) is None:
                    break
            else:
                if current:
                    logger.debug(f"Found attribute via nested path: {path}")
                    return current
        # Simple attribute access
        elif hasattr(obj, path) and (val := getattr(obj, path)):
            logger.debug(f"Found attribute via direct access: {path}")
            return val
    
    # Check properties dictionary
    if hasattr(obj, 'properties'):
        props = obj.properties
        for path in paths:
            if path in props and props[path]:
                logger.debug(f"Found attribute via properties: {path}")
                return props[path]
    
    # Try JSON serialization as last resort
    if hasattr(obj, 'to_json'):
        data_dict = json.loads(obj.to_json())
        for path in paths:
            if path in data_dict and data_dict[path]:
                logger.debug(f"Found attribute via JSON: {path}")
                return data_dict[path]
    
    return default

def get_sender_info(message: Any) -> Tuple[str, str]:
    """Extract sender name and email from a message"""
    # Common paths for sender information
    email_paths = [
        'sender.emailAddress.address', 'from_.emailAddress.address',
        'from.emailAddress.address', 'sender.email_address.address',
        'from_.email_address.address'
    ]
    name_paths = [
        'sender.emailAddress.name', 'from_.emailAddress.name', 
        'from.emailAddress.name', 'sender.email_address.name',
        'from_.email_address.name'
    ]
    
    # Try getting the email and matching name
    sender_email = get_message_attribute(message, email_paths, '')
    sender_name = get_message_attribute(message, name_paths, '') if sender_email else ""
    
    return sender_name, sender_email

def get_date(message: Any) -> Any:
    """Extract date from a message with formatting"""
    # Common date paths
    date_paths = [
        'receivedDateTime', 'received_date_time', 'sentDateTime', 'sent_date_time', 
        'createdDateTime', 'created_date_time', 'lastModifiedDateTime', 'last_modified_date_time'
    ]
    
    date_value = get_message_attribute(message, date_paths)
    
    # Format the date if found
    if not date_value:
        return "Unknown"
    
    try:
        if isinstance(date_value, str):
            if 'T' in date_value:
                date_value = datetime.fromisoformat(date_value.replace('Z', '+00:00'))
            else:
                for fmt in ['%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%d/%m/%Y %H:%M:%S']:
                    try:
                        date_value = datetime.strptime(date_value, fmt)
                        break
                    except ValueError:
                        continue
        
        if isinstance(date_value, datetime):
            return date_value.strftime('%Y-%m-%d %H:%M:%S')
        
        return str(date_value)
    except Exception as e:
        logger.warning(f"Error formatting date: {e}")
        return str(date_value)

@safe_operation()
def format_email_output(message: Any, skip_cleaning: bool = False, as_text=False) -> Any:
    """Formats an email by cleaning its HTML content"""
    if as_text:
        return format_email_as_text(message)
    
    if skip_cleaning or not message:
        return message
    
    # Clean HTML content if present
    if hasattr(message, 'body') and message.body:
        body = message.body
        content = get_message_attribute(body, ['content'], '')
        content_type = get_message_attribute(body, ['content_type', 'contentType'], '').lower()
        
        if content_type == 'html' and content:
            clean_text = clean_email_content(content)
            
            # Update the body content
            if hasattr(body, 'set_property'):
                body.set_property('content', clean_text)
                body.set_property('contentType', 'text')
            elif hasattr(body, 'content'):
                body.content = clean_text
                if hasattr(body, 'contentType'):
                    body.contentType = 'text'
    
    return message

@safe_operation("")
def format_email_as_text(message: Any) -> str:
    """Creates a plain text representation of an email"""
    if not message:
        return "[No message data]"
    
    # Extract basic message information
    subject = get_message_attribute(message, ['subject'], 'No Subject')
    sender_name, sender_email = get_sender_info(message)
    formatted_date = get_date(message)
    message_id = get_message_attribute(message, ['id'], 'Unknown ID')
    
    # Build the formatted output
    lines = [
        f"EMAIL: {subject}",
        f"From: {f'{sender_name} <{sender_email}>' if sender_name else sender_email or 'Unknown'}",
        f"Date: {formatted_date}",
        f"ID: {message_id}"
    ]
    
    # Extract and format content
    content = ""
    content_type = "text"
    
    if hasattr(message, 'body') and message.body:
        content = get_message_attribute(message.body, ['content'], '')
        content_type = get_message_attribute(message.body, ['content_type', 'contentType'], 'text').lower()
    
    lines.append(f"Type: {content_type}")
    lines.append("-" * 50)
    
    if content:
        if content_type == 'html':
            lines.append(clean_email_content(content))
        else:
            lines.append(re.sub(r'\n{3,}', '\n\n', 
                      re.sub(r'[\x00-\x09\x0B\x0C\x0E-\x1F\x7F]', '', content)))
    else:
        lines.append("[No Content]")
    
    return '\n'.join(lines)