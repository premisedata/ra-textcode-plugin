# Python functions ported from db-api.py for Pyodide
# These are loaded into Pyodide at runtime via taskpane.js

import re
import json
from datetime import datetime

def parseParticipantNumber(tagged_text, paragraph):
    """Parse participant number from text (P1:, P2:, etc.)"""
    ret = None
    match = re.search(r"([P_]\d+):", tagged_text)
    if match:
        ret = match.group(1).replace('_','P')
    else:
        match = re.search(r"([P_]\d+):", paragraph)
        if match:
            ret = match.group(1).replace('_','P')
    return ret

def getProjectName(doc_filename):
    """Determine project name from filename"""
    ret = ""
    if "_FG" in doc_filename:
        ret = "DIANA"
    if "_LVA_" in doc_filename:
        ret = "JOHNSON"
    if "RIDE_" in doc_filename:
        ret = "RIDE"
    if "JOHNSON2_" in doc_filename:
        ret = "JOHNSON2"
    return ret

def parseCountryName(doc_filename):
    """Parse country name from DIANA filename format"""
    ret = '<NO-COUNTRY-FOUND>'
    a = doc_filename.split('_FG')
    if len(a) >= 2:
        b = a[0].split('_')
        if len(b) == 2:
            ret = b[1]
    return ret

def cleanCategory(c):
    """Remove display_order suffix from category string"""
    ret = c.replace(' -- display_order=None','')
    ret = re.sub(r" -- display_order=\d+", "", ret)
    return ret

def createCitationJOHNSON(citation):
    """Create citation string for JOHNSON project"""
    gender = ""
    age = ""
    occupation = ""
    city = ""
    res = re.search(r"<gender>(.*?)</gender>", citation)
    if res:
        gender = res.group(1)
    res = re.search(r"<age>(.*?)</age>", citation)
    if res:
        age = res.group(1)
    res = re.search(r"<occupation>(.*?)</occupation>", citation)
    if res:
        occupation = res.group(1)
    res = re.search(r"<city>(.*?)</city>", citation)
    if res:
        city = res.group(1)
    return f"({gender}, {age}, {occupation}, {city})"

def createCitationJOHNSON2(citation, participant_number):
    """Create citation string for JOHNSON2 project"""
    gender = ""
    age = ""
    occupation = ""
    city = ""
    res = re.search(r"<gender>(.*?)</gender>", citation)
    if res:
        gender = res.group(1)
    res = re.search(r"<age>(.*?)</age>", citation)
    if res:
        age = res.group(1)
    res = re.search(r"<occupation>(.*?)</occupation>", citation)
    if res:
        occupation = res.group(1)
    res = re.search(r"<city>(.*?)</city>", citation)
    if res:
        city = res.group(1)
    return f"({gender}, {age}, {occupation}, {city})"

def createCitationRIDE(citation, doc_filename):
    """Create citation string for RIDE project"""
    gender = ""
    age = ""
    occupation = ""
    city = ""
    res = re.search(r"<gender>(.*?)</gender>", citation)
    if res:
        gender = res.group(1)
    res = re.search(r"<age>(.*?)</age>", citation)
    if res:
        age = res.group(1)
    
    occupation = 'Unk'
    if 'former_member' in doc_filename.lower():
        occupation = 'Former Member'
    elif 'recruited_non_member' in doc_filename.lower():
        occupation = 'Recruited Non-Member'
    elif 'current_member' in doc_filename.lower():
        occupation = 'Current Member'
    
    res = re.search(r"<city>(.*?)</city>", citation)
    if res:
        city = res.group(1)
    
    return f"({gender}, {age}, {occupation}, {city})"

