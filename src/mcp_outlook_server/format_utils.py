import re, json, logging
from datetime import datetime
from functools import wraps
from typing import Any, Tuple

logger = logging.getLogger(__name__)

def safe_operation(default_return=None):
    # Handle exceptions in operations.
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try: return func(*args, **kwargs)
            except Exception as e:
                logger.debug(f"Error in {func.__name__}: {e}")
                return default_return
        return wrapper
    return decorator

@safe_operation()
def get_message_attribute(obj: Any, paths: list, default=None) -> Any:
    # Get attributes from message objects.
    if not obj: return default
    
    if hasattr(obj, 'to_json'): # Try JSON serialization for Microsoft Graph objects.
        try:
            data_dict = json.loads(obj.to_json())
            for path in paths:
                if path in data_dict and data_dict[path]: return data_dict[path]
        except Exception as e: logger.debug(f"Error in JSON serialization: {e}")
    
    for path in paths: # Fast path for direct attribute access.
        if '.' in path:
            parts, current = path.split('.'), obj
            try:
                for part in parts: current = getattr(current, part)
                if current is not None: return current
            except AttributeError: continue
        elif hasattr(obj, path) and (val := getattr(obj, path)) is not None: return val
    
    if hasattr(obj, 'properties') and (props := obj.properties): # Check properties dictionary.
        for path in paths:
            if path in props and props[path]: return props[path]
    return default

def get_sender_info(message: Any) -> Tuple[str, str]:
    # Extract sender name and email.
    email_paths = ['sender.emailAddress.address', 'from_.emailAddress.address', 'from.emailAddress.address']
    name_paths = ['sender.emailAddress.name', 'from_.emailAddress.name', 'from.emailAddress.name']
    sender_email = get_message_attribute(message, email_paths, '')
    sender_name = get_message_attribute(message, name_paths, '') if sender_email else ""
    return sender_name, sender_email

def format_date_string(date_value: Any) -> str:
    # Format date to string.
    if not date_value: return "Unknown"
    try:
        if isinstance(date_value, str):
            dt_obj = None
            if 'T' in date_value: dt_obj = datetime.fromisoformat(date_value.replace('Z', '+00:00'))
            else:
                for fmt in ('%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%d/%m/%Y %H:%M:%S'):
                    try: dt_obj = datetime.strptime(date_value, fmt); break
                    except ValueError: continue
            date_value = dt_obj if dt_obj else date_value
        return date_value.strftime('%Y-%m-%d %H:%M:%S') if isinstance(date_value, datetime) else str(date_value)
    except Exception as e:
        logger.warning(f"Error formatting date '{date_value}': {e}")
        return str(date_value)

def get_date(message: Any) -> str:
    # Extract and format date from message.
    date_paths_prop = ['receivedDateTime', 'sentDateTime', 'createdDateTime', 'lastModifiedDateTime']
    date_paths_attr = ['received_datetime', 'sent_datetime', 'created_datetime', 'last_modified_datetime']
    
    date_val = None
    if hasattr(message, 'properties') and message.properties:
        for path in date_paths_prop:
            if path in message.properties and message.properties[path]:
                date_val = message.properties[path]; break
    if not date_val: date_val = get_message_attribute(message, date_paths_attr)
    
    return format_date_string(date_val) if date_val else "Unknown"

@safe_operation()
def format_email_output(
    message: Any, skip_cleaning: bool = False, as_text: bool = False, deep_clean: bool = True
) -> Any:
    """Normalise the email body or return as plain text."""

    if as_text:
        return format_email_as_text(message, deep_clean=deep_clean)
    if skip_cleaning or not message or not hasattr(message, "body") or not message.body:
        return message

    body = message.body
    content = get_message_attribute(body, ["content"], "")
    content_type = get_message_attribute(body, ["content_type", "contentType"], "").lower()

    if content_type == "html" and content:
        from .clean_utils import clean_email_content  # Local import to avoid circularity.

        clean_content = clean_email_content(content, deep_clean=deep_clean)
        if hasattr(body, "set_property"):
            body.set_property("content", clean_content)
            body.set_property("contentType", "text")
        elif hasattr(body, "content"):
            body.content = clean_content
            if hasattr(body, "contentType"):
                body.contentType = "text"
    return message

@safe_operation("")
def format_email_as_text(message: Any, deep_clean: bool = True) -> str:
    """Render an email object as a plain text string."""

    if not message:
        return "[No message data]"

    subject = get_message_attribute(message, ["subject"], "No Subject")
    sender_name, sender_email = get_sender_info(message)
    sender_display = (
        f"{sender_name} <{sender_email}>" if sender_name else sender_email or "Unknown"
    )

    lines = [
        f"EMAIL: {subject}",
        f"From: {sender_display}",
        f"Date: {get_date(message)}",
        f"ID: {get_message_attribute(message, ['id'], 'Unknown ID')}",
    ]

    content, content_type = "", "text"
    if hasattr(message, "body") and message.body:
        content = get_message_attribute(message.body, ["content"], "")
        content_type = get_message_attribute(
            message.body, ["content_type", "contentType"], "text"
        ).lower()

    lines.extend([f"Type: {content_type}", "-" * 50])

    if content:
        if content_type == "html":
            from .clean_utils import clean_email_content  # Local import.

            lines.append(clean_email_content(content, deep_clean=deep_clean))
        else:  # Basic cleaning for non-HTML plain text.
            cleaned_plain = re.sub(
                r"[\x00-\x09\x0B\x0C\x0E-\x1F\x7F]", "", content or ""
            )
            lines.append(re.sub(r"\n{3,}", "\n\n", cleaned_plain))
    else:
        lines.append("[No Content]")

    return "\n".join(lines)
