# format_utils_search.py
"""
Formateo para resultados de búsqueda $search (KQL) de Microsoft Graph API (JSON plano).
"""
from typing import Any, Dict, List
from .clean_utils import clean_email_content, clean_text, process_text

def format_email_output_search(
    msg: Dict[str, Any], as_text: bool = True, clean_body: bool = True
) -> Dict[str, Any]:
    """Format plain JSON Graph API message for uniform output."""

    body_raw = msg.get("body", {}).get("content") if msg.get("body") else None
    content_type = (
        msg.get("body", {}).get("contentType", "text").lower()
        if msg.get("body")
        else "text"
    )
    if body_raw and content_type == "html":
        body_clean = clean_email_content(body_raw, aggressive=True, deep_clean=clean_body)
    elif body_raw:
        body_clean = process_text(clean_text(body_raw, aggressive=True))
    else:
        body_clean = None
    return {
        "id": msg.get("id", "Desconocido"),
        "subject": msg.get("subject", "Sin asunto"),
        "from": _format_address(msg.get("from")),
        "toRecipients": [_format_address(r) for r in msg.get("toRecipients", [])],
        "ccRecipients": [_format_address(r) for r in msg.get("ccRecipients", [])],
        "bccRecipients": [_format_address(r) for r in msg.get("bccRecipients", [])],
        "sentDateTime": msg.get("sentDateTime"),
        "receivedDateTime": msg.get("receivedDateTime"),
        "isRead": msg.get("isRead"),
        "importance": msg.get("importance"),
        "categories": msg.get("categories", []),
        "bodyPreview": msg.get("bodyPreview"),
        "body": body_clean,
    }

def format_email_output_search_no_body(msg: Dict[str, Any], as_text: bool = True) -> Dict[str, Any]:
    # Formatea un mensaje JSON plano de Graph API para salida homogénea sin el cuerpo
    return {
        "id": msg.get("id", "Desconocido"),
        "subject": msg.get("subject", "Sin asunto"),
        "from": _format_address(msg.get("from")),
        "toRecipients": [_format_address(r) for r in msg.get("toRecipients", [])],
        "ccRecipients": [_format_address(r) for r in msg.get("ccRecipients", [])],
        "bccRecipients": [_format_address(r) for r in msg.get("bccRecipients", [])],
        "sentDateTime": msg.get("sentDateTime"),
        "receivedDateTime": msg.get("receivedDateTime"),
        "isRead": msg.get("isRead"),
        "importance": msg.get("importance"),
        "categories": msg.get("categories", []),
        "bodyPreview": msg.get("bodyPreview")
        # No body
    }

def _format_address(addr: Any) -> str:
    if not addr:
        return ""
    if isinstance(addr, dict) and "emailAddress" in addr:
        a = addr["emailAddress"]
        return f'{a.get("name", "")} <{a.get("address", "")}>' if a.get("address") else a.get("name", "")
    if isinstance(addr, dict) and "address" in addr:
        return addr.get("address", "")
    return str(addr)

def format_emails_list_structured_search(
    messages: List[Dict[str, Any]], clean_body: bool = True
) -> List[Dict[str, Any]]:
    return [format_email_output_search(msg, clean_body=clean_body) for msg in messages]

def format_emails_list_structured_search_no_body(
    messages: List[Dict[str, Any]], **_
) -> List[Dict[str, Any]]:
    return [format_email_output_search_no_body(msg) for msg in messages]

