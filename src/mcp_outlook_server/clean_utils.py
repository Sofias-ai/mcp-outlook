import re, logging, html2text, unicodedata
from bs4 import BeautifulSoup
from functools import wraps

# Configure logging
logger = logging.getLogger(__name__)

def exception_handler(default_return=""):
    """Decorator to handle exceptions in text processing functions"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                logger.error(f"Error in {func.__name__}: {e}")
                return default_return if args[0] is None else args[0]
        return wrapper
    return decorator

@exception_handler()
def clean_text(text, aggressive=False):
    """Cleans and normalizes text"""
    if not text: return ""
    
    # Remove BOM, zero-width, and control characters
    text = re.sub(r'[\ufeff\u200b-\u200f\u2028\u2029\u202a-\u202e\x00-\x09\x0B\x0C\x0E-\x1F\x7F]', '', text)
    
    # Additional aggressive cleaning
    if aggressive:
        text = re.sub(r'[\u2060-\u2064\u206A-\u206F\u00AD\u034F]', '', text)  # Remove invisible chars
        text = re.sub(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])', '', text)     # Remove ANSI codes
    
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'^ +| +$', '', text, flags=re.MULTILINE)
    
    # Normalize Unicode
    return unicodedata.normalize('NFC', text)

@exception_handler()
def html_to_text(html_content):
    """Converts HTML to plain text with improved formatting"""
    if not html_content: return ""
    
    # Parse HTML
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Clean up unwanted elements
    for tag in soup(['script', 'style']):
        tag.decompose()
    
    # Replace images with placeholder
    for img in soup.find_all('img'):
        img.replace_with('[Image]')
    
    # Add spacing around tables
    for table in soup.find_all('table'):
        table.insert_before('\n')
        table.insert_after('\n')
    
    # Convert to text
    h = html2text.HTML2Text()
    h.ignore_links = False
    h.body_width = 0
    h.images_to_alt = True
    h.protect_links = True
    
    text = h.handle(str(soup))
    
    # Fix common formatting issues
    text = re.sub(r'!\[\]\([^)]+\)', '[Image]', text)
    text = re.sub(r'\[Image\]\(<\[Link\]\)', '[Image]', text)
    text = re.sub(r'\[([^]]+)\]\(<\[Link\]\)', r'[\1]', text)
    text = re.sub(r'\[([^]]+)\]\(mailto:[^)]+\)', r'[\1]', text)
    
    # Clean up excessive formatting
    text = re.sub(r'---(\s*\|?\s*---)*', '---', text)
    text = re.sub(r'(\n\s*[-]+\s*\n)+', '\n\n', text)
    text = re.sub(r'\n\s*\|\s*\n', '\n\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    return text

@exception_handler()
def process_text(text):
    """Combined function for text processing operations"""
    if not text:
        return ""
    
    # Remove URLs and references
    url_patterns = [
        r'https?://\S+', r'ftp://\S+', r'www\.\S+\.[a-zA-Z]{2,}(/\S*)?',
        r'mailto:\S+', r'tel:[+\d\s\(\)-]+',
        r'[a-zA-Z0-9_\-\.]+\.(com|org|net|io|ai|app|co|edu|gov|mil|biz|info|me|tv|xyz|eu|uk|es|de|fr)(/\S*)?'
    ]
    for pattern in url_patterns:
        text = re.sub(pattern, '[Link]', text)
    
    # Remove duplicate lines
    lines = text.split('\n')
    unique_lines = []
    prev_line = None
    for line in lines:
        if not re.match(r'^\s*[\|\-]+\s*$', line) and line != prev_line:
            unique_lines.append(line)
        prev_line = line
    
    # Look for disclaimers and trim content
    result = '\n'.join(unique_lines)
    disclaimer_match = re.search(r'(?i)(confidential|disclaimer|privileged|legal notice|aviso legal|email\s+disclaimer|this email and any files|this message and any attachments)', result)
    if disclaimer_match:
        result = result[:disclaimer_match.start()].strip()
    
    return result

def clean_email_content(html_content, aggressive=True):
    """Main pipeline for cleaning email content"""
    if not html_content: return ""
    
    try:
        # Apply processing pipeline
        text = clean_text(html_content, aggressive)
        text = html_to_text(text)
        text = process_text(text)
        return re.sub(r'\n{3,}', '\n\n', text)
    except Exception as e:
        logger.error(f"Error in email cleaning pipeline: {e}")
        # Fallback cleanup
        return re.sub(r'<[^>]+>', ' ', html_content).strip()