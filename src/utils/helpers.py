"""
Utility helper functions
"""
import re
from urllib.parse import urlparse, urlencode, parse_qs


def extract_year_from_text(text):
    """Extract year from text using various patterns"""
    patterns = [
        r'(?:20)(1[5-9]|2[0-9])\b',  # 2015-2029
        r'\b(20\d{2})\b',  # 20xx
        r'(?:3112|1312|31_12_|31-12-)(1[5-9]|2[0-9])',  # 311219 or 31_12_19
        r'(?:3112|1312|31_12_|31-12-)(20\d{2})',  # 31122019
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            year = match.group(1) if match.lastindex and match.lastindex >= 1 else match.group(0)
            year = re.sub(r'[^0-9]', '', year)
            
            if len(year) == 2:
                year = f"20{year}"
            
            if len(year) == 4 and 2010 <= int(year) <= 2030:
                return year
    return None


def extract_trailing_numbers(text):
    """
    Extracts numerical values from the end of a string.
    Example: "Capital social 50 000 000" -> ("Capital social", 50000000)
    """
    if not text:
        return text, []

    # Match trailing group of numbers (including spaces as thousands separators)
    # We look for groups of digits potentially separated by spaces or dots
    # e.g., "123 456" or "123.456" or "123"
    pattern = r'\s+([\d\s.,]+)$'
    match = re.search(pattern, text)
    
    if match:
        potential_num_str = match.group(1).strip()
        # Clean the number string
        cleaned_num = re.sub(r'[\s.]', '', potential_num_str).replace(',', '.')
        
        # Verify if it's actually a number
        try:
            val = float(cleaned_num)
            # If successful, return the text without the number and the number itself
            remaining_text = text[:match.start()].strip()
            return remaining_text, [val]
        except ValueError:
            pass
            
    return text, []


def clean_number(text):
    """Clean and convert text to number"""
    if isinstance(text, str):
        cleaned = re.sub(r'\s+', '', text).replace(',', '.').rstrip('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ')
        try:
            return int(float(cleaned))
        except ValueError:
            return text
    return int(text) if isinstance(text, (int, float)) else text


def normalize_url(url):
    """Normalize URL by removing dynamic parameters"""
    parsed_url = urlparse(url)
    query_params = parse_qs(parsed_url.query)
    if 'id' in query_params or 'token' in query_params:
        query_params.pop('id', None)
        query_params.pop('token', None)
    new_query = urlencode(query_params, doseq=True)
    return parsed_url._replace(query=new_query).geturl()
