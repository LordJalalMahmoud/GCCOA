import re
import json
import os
from typing import Dict, Any, Optional, Tuple, List

# Load Bosta locations
LOCATIONS_FILE = os.path.join(os.path.dirname(__file__), 'bosta_locations.json')
BOSTA_CITIES: List[str] = []
BOSTA_CITY_AREAS: Dict[str, List[str]] = {}
AREA_TO_CITY_MAP: Dict[str, str] = {}

if os.path.exists(LOCATIONS_FILE):
    with open(LOCATIONS_FILE, 'r', encoding='utf-8') as f:
        BOSTA_CITY_AREAS = json.load(f)
        BOSTA_CITIES = list(BOSTA_CITY_AREAS.keys())
        for city, areas in BOSTA_CITY_AREAS.items():
            for area in areas:
                norm_area = area.strip()
                if norm_area:
                    AREA_TO_CITY_MAP[norm_area.lower()] = city

# Common aliases / normalizations for Egyptian cities
CITY_ALIASES = {
    'القاهره': 'القاهرة',
    'cairo': 'القاهرة',
    'الجيزه': 'الجيزة',
    'giza': 'الجيزة',
    'اسكندرية': 'الإسكندرية',
    'الاسكندرية': 'الإسكندرية',
    'الاسكندريه': 'الإسكندرية',
    'alexandria': 'الإسكندرية',
    'alex': 'الإسكندرية',
    'الشرقيه': 'الشرقية',
    'sharqia': 'الشرقية',
    'الدقهليه': 'الدقهلية',
    'daqahlia': 'الدقهلية',
    'المنوفيه': 'المنوفية',
    'menofia': 'المنوفية',
    'الغربيه': 'الغربية',
    'gharbia': 'الغربية',
    'القليوبيه': 'القليوبية',
    'qalyubia': 'القليوبية',
    'البحيره': 'البحيرة',
    'beheira': 'البحيرة',
    'الاسماعيليه': 'الإسماعيلية',
    'ismailia': 'الإسماعيلية',
    'بورسعيد': 'بور سعيد',
    'بورفؤاد': 'بور سعيد',
    'port said': 'بور سعيد',
    'كفرالشيخ': 'كفر الشيخ',
    'بنى سويف': 'بني سويف',
    'بني سويف': 'بني سويف',
    'bni suef': 'بني سويف',
    'بني سويف ': 'بني سويف',
    'الفيوم': 'الفيوم',
    'fayoum': 'الفيوم',
    'المنيا': 'المنيا',
    'minya': 'المنيا',
    'اسيوط': 'أسيوط',
    'assiut': 'أسيوط',
    'سوهاج': 'سوهاج',
    'sohag': 'سوهاج',
    'قنا': 'قنا',
    'qena': 'قنا',
    'الاقصر': 'الأقصر',
    'luxor': 'الأقصر',
    'اسوان': 'أسوان',
    'aswan': 'أسوان',
    'السويس': 'السويس',
    'suez': 'السويس',
    'دمياط': 'دمياط',
    'damietta': 'دمياط',
    'البحر الاحمر': 'البحر الأحمر',
    'red sea': 'البحر الأحمر',
    'الغردقة': 'البحر الأحمر',
    'الغردقه': 'البحر الأحمر',
    'مرسى مطروح': 'مرسى مطروح',
    'مطروح': 'مرسى مطروح',
    'matrouh': 'مرسى مطروح',
    'الساحل': 'الساحل الشمالي',
    'الساحل الشمالي': 'الساحل الشمالي',
    'شمال سيناء': 'جنوب سيناء',
    'جنوب سيناء': 'جنوب سيناء',
    'شرم الشيخ': 'جنوب سيناء',
}

# Arabic-Indic to standard ASCII digits
ARABIC_DIGITS_TABLE = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')

def normalize_digits(text: str) -> str:
    """Converts Arabic-Indic digits to Latin digits."""
    if not text:
        return ""
    return str(text).translate(ARABIC_DIGITS_TABLE)

def clean_phone(raw: str) -> str:
    """Cleans phone numbers: converts digits, strips spaces, symbols, and +20 country code."""
    if not raw:
        return ""
    digits = normalize_digits(raw)
    digits = re.sub(r'[^\d]', '', digits)
    # Strip country code 20 if it leads to 12 or 13 digits (e.g. 002010... or 2010...)
    if digits.startswith('0020'):
        digits = digits[4:]
    elif digits.startswith('20') and len(digits) >= 12:
        digits = digits[2:]
    return digits

def validate_egyptian_phone(phone: str) -> Tuple[bool, str, str]:
    """
    Validates Egyptian phone number.
    Returns (is_valid, normalized_phone, error_message).
    Egyptian mobiles: 11 digits starting with 010, 011, 012, or 015.
    """
    cleaned = clean_phone(phone)
    if not cleaned:
        return False, "", "رقم الهاتف فارغ!"

    # If user provided 10 digits starting with 10, 11, 12, 15 (missing leading 0)
    if len(cleaned) == 10 and cleaned.startswith(('10', '11', '12', '15')):
        cleaned = '0' + cleaned
        return True, cleaned, "تم إضافة الصفر في بداية الرقم تلقائياً."

    if len(cleaned) < 11:
        return False, cleaned, f"رقم الهاتف ناقص! عدد الأرقام {len(cleaned)} بدلاً من 11 رقم ({cleaned})"

    if len(cleaned) > 11:
        return False, cleaned, f"رقم الهاتف زائد! عدد الأرقام {len(cleaned)} بدلاً من 11 رقم ({cleaned})"

    if not re.match(r'^01[0125]\d{8}$', cleaned):
        return False, cleaned, f"رقم الهاتف غير صالح ({cleaned}): يجب أن يبدأ بـ 010 أو 011 أو 012 أو 015."

    return True, cleaned, ""


def find_phones_in_text(text: str) -> List[str]:
    """Finds all potential phone numbers in text."""
    norm_text = normalize_digits(text)
    # Patterns for egyptian phone numbers, including possible separators
    # E.g., 010 1234 5678, +2010..., 010-123-4567, or 10 digits
    matches = re.findall(r'(?:\+?20|0020)?[\s-]*(?:0?1[0125][\s-]*\d{4}[\s-]*\d{4}|\b01[0125]\d{8}\b|\b1[0125]\d{8}\b|\b0\d{8,10}\b)', norm_text)
    
    cleaned_phones = []
    for m in matches:
        cp = clean_phone(m)
        if len(cp) == 10 and cp.startswith(('10', '11', '12', '15')):
            cp = '0' + cp
        if cp and cp not in cleaned_phones:
            cleaned_phones.append(cp)
    
    # Also search for standard sequence of 10 or 11 digits if matches is empty
    if not cleaned_phones:
        any_num = re.findall(r'\b\d{10,11}\b', norm_text)
        for num in any_num:
            cp = clean_phone(num)
            if cp not in cleaned_phones:
                cleaned_phones.append(cp)

    return cleaned_phones


def match_city_and_area(text: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Intelligently extracts City and Area from text using Bosta's location dictionary.
    """
    clean_text = text.lower()
    
    found_city = None
    found_area = None
    
    # 1. First check explicit city aliases
    for alias, city in CITY_ALIASES.items():
        # Match as whole word or with prefix
        pattern = r'(?:\b|\s|^|،|-)' + re.escape(alias) + r'(?:\b|\s|$|،|-)'
        if re.search(pattern, clean_text):
            found_city = city
            break
            
    # 2. Check if text mentions an exact Bosta city
    if not found_city:
        for city in BOSTA_CITIES:
            pattern = r'(?:\b|\s|^|،|-)' + re.escape(city.lower()) + r'(?:\b|\s|$|،|-)'
            if re.search(pattern, clean_text):
                found_city = city
                break
                
    # 3. Check for specific areas
    # Sort areas by length descending so longer specific names match first (e.g. 'الشيخ زايد' before 'زايد')
    all_areas_sorted = sorted(AREA_TO_CITY_MAP.keys(), key=len, reverse=True)
    
    # If city is already known, prioritize areas belonging to that city
    if found_city and found_city in BOSTA_CITY_AREAS:
        city_areas = sorted(BOSTA_CITY_AREAS[found_city], key=len, reverse=True)
        for area in city_areas:
            area_lower = area.lower()
            if area_lower in clean_text:
                found_area = area
                break
                
    # If no area found yet, search all areas
    if not found_area:
        for area_lower in all_areas_sorted:
            if len(area_lower) < 3:
                continue
            # Search area with boundaries
            pattern = r'(?:\b|\s|^|،|-)' + re.escape(area_lower) + r'(?:\b|\s|$|،|-)'
            if re.search(pattern, clean_text):
                found_area = area_lower
                # If city wasn't found, deduce city from area!
                if not found_city:
                    found_city = AREA_TO_CITY_MAP.get(area_lower)
                break

    # Format found_area to match official casing/spelling from city_areas if possible
    if found_city and found_area:
        for official_area in BOSTA_CITY_AREAS.get(found_city, []):
            if official_area.lower() == found_area.lower():
                found_area = official_area
                break

    return found_city, found_area


def clean_product_text(text: str) -> str:
    """
    Strips any order codes, tracking codes, SKU numbers, or reference tags from product description.
    Ensures Column N (وصف الشحنة) contains ONLY the products and their quantities.
    """
    if not text:
        return ""
    cleaned = str(text).strip()
    # 1. Remove bracketed codes like (كود GC-02) or (GC01) or [ORD-12] or (كود: 123)
    cleaned = re.sub(r'[\(\[\{]\s*(?:كود(?:\s+الطلب|\s+الشحن|\s+المنتج)?|مرجع|المرجع|رمز|SKU|Code|Ref)?\s*[:=\-#]?\s*[\w\-]+\s*[\)\]\}]', '', cleaned, flags=re.IGNORECASE)
    # 2. Remove explicit code labels and their values (e.g. كود: GC01, كود الشحن: 12345, كود الطلب: ORD-9)
    cleaned = re.sub(r'(?:كود(?:\s+الطلب|\s+الشحن|\s+المنتج)?|مرجع|المرجع|رمز|SKU|Code|Ref)\s*[:=\-#]?\s*[\w\-]+', '', cleaned, flags=re.IGNORECASE)
    # 3. Remove standalone SKU patterns, ORD codes, or hash tags
    cleaned = re.sub(r'\bORD[-_]?\d+\b', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\b[A-Za-z]{1,4}[-_]?\d{1,6}\b', '', cleaned)
    cleaned = re.sub(r'#\w+', '', cleaned)
    # 4. Remove empty brackets
    cleaned = re.sub(r'[\(\[\{]\s*[\)\]\}]', '', cleaned)
    # 5. Clean up redundant spaces and separators
    cleaned = re.sub(r'\s*([,\-،_])\s*', r' \1 ', cleaned)
    cleaned = re.sub(r'(\s*[-_،,]\s*){2,}', ' - ', cleaned)
    cleaned = re.sub(r'^\s*[-_:,،]\s*', '', cleaned)
    cleaned = re.sub(r'\s*[-_:,،]\s*$', '', cleaned)
    cleaned = ' '.join(cleaned.split()).strip()
    return cleaned


def parse_order_text(text: str) -> Dict[str, Any]:
    """
    Ultra-fast rule & pattern-based order extractor.
    Parses unstructured and structured Arabic/English order text.
    """
    text = normalize_digits(text).strip()
    result = {
        "name": "",
        "phone": "",
        "second_phone": "",
        "city": "",
        "area": "",
        "address": "",
        "landmark": "",
        "cod_amount": "",
        "items_count": "1",
        "description": "",
        "notes": "",
        "shipment_value": "",
        "raw_text": text,
        "warnings": [],
        "errors": []
    }
    
    if not text:
        return result

    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    # Extract phones first
    all_phones = find_phones_in_text(text)
    if all_phones:
        result["phone"] = all_phones[0]
        if len(all_phones) > 1:
            result["second_phone"] = all_phones[1]

    # Keyword-based line extraction
    unhandled_lines = []
    
    # Regex patterns for fields
    NAME_PATTERNS = [
        r'^(?:الاسم|اسم العميل|اسم المستلم|العميل|Name)\s*[:=\-]\s*(.+)$',
        r'^(?:الاسم|اسم|Name)\s+(.+)$'
    ]
    PHONE_PATTERNS = [
        r'^(?:الموبايل|التليفون|الهاتف|رقم الموبايل|رقم التليفون|رقم الهاتف|الرقم|فون|تليفون|موبايل|Phone|Mobile)\s*[:=\-]\s*(.+)$'
    ]
    SECOND_PHONE_PATTERNS = [
        r'^(?:رقم ثاني|رقم اخر|تليفون ثاني|هاتف ارضي|الرقم الثاني|رقم تاني|Second Phone)\s*[:=\-]\s*(.+)$'
    ]
    CITY_PATTERNS = [
        r'^(?:المحافظة|المدينة|المحافظه|المدينه|City|Gov)\s*[:=\-]\s*(.+)$'
    ]
    AREA_PATTERNS = [
        r'^(?:المنطقة|المنطقه|الحي|Area|District)\s*[:=\-]\s*(.+)$'
    ]
    ADDRESS_PATTERNS = [
        r'^(?:العنوان|عنوان العميل|مكان التوصيل|الشارع|العنوان بالتفصيل|Address)\s*[:=\-]\s*(.+)$'
    ]
    PRICE_PATTERNS = [
        r'^(?:السعر|المبلغ|قيمة التحصيل|التحصيل|المطلوب|الاجمالي|الإجمالي|الحساب|سعر الأوردر|سعر الاوردر|Price|COD|Total|Amount)\s*[:=\-]\s*(.+)$'
    ]
    ITEMS_PATTERNS = [
        r'^(?:القطع|عدد القطع|الكمية|الكميه|عدد|Items|Quantity|Qty)\s*[:=\-]\s*(.+)$'
    ]
    DESC_PATTERNS = [
        r'^(?:المنتجات|تفاصيل منتجات الطلب|الأصناف|الاصناف|محتويات الشحنة|محتوى الشحنة|الطلبية|الطلبيه|الوصف|الاوردر|الأوردر|الطلب|المنتج|المحتوى|الشحنة|وصف الشحنة|تفاصيل الاوردر|تفاصيل الأوردر|Order|Description|Package|Products)\s*[:=\-]\s*(.+)$'
    ]
    NOTES_PATTERNS = [
        r'^(?:ملاحظات التوصيل|ملاحظات الشحن|تعليمات التوصيل|ملاحظات|الملاحظات|ملحوظة|ملاحظة|Notes|Note|Delivery Notes)\s*[:=\-]\s*(.+)$'
    ]

    for line in lines:
        matched = False
        
        # Check Name
        for pat in NAME_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                val = m.group(1).strip()
                # filter out if user just wrote phone or address
                if not re.match(r'^\d+$', val):
                    result["name"] = val
                    matched = True
                    break
        if matched: continue
        
        # Check Second Phone
        for pat in SECOND_PHONE_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                phones = find_phones_in_text(m.group(1))
                if phones:
                    result["second_phone"] = phones[0]
                else:
                    result["second_phone"] = clean_phone(m.group(1))
                matched = True
                break
        if matched: continue

        # Check Primary Phone
        for pat in PHONE_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                phones = find_phones_in_text(m.group(1))
                if phones:
                    result["phone"] = phones[0]
                matched = True
                break
        if matched: continue

        # Check City
        for pat in CITY_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                c_val = m.group(1).strip()
                city, area = match_city_and_area(c_val)
                if city:
                    result["city"] = city
                else:
                    result["city"] = c_val
                if area and not result["area"]:
                    result["area"] = area
                matched = True
                break
        if matched: continue

        # Check Area
        for pat in AREA_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                a_val = m.group(1).strip()
                city, area = match_city_and_area(a_val)
                result["area"] = area or a_val
                if city and not result["city"]:
                    result["city"] = city
                matched = True
                break
        if matched: continue

        # Check Address
        for pat in ADDRESS_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                result["address"] = m.group(1).strip()
                matched = True
                break
        if matched: continue

        # Check Price / COD
        for pat in PRICE_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                val = m.group(1).strip()
                # extract numeric price
                num_m = re.search(r'(\d+(?:\.\d+)?)', val)
                if num_m:
                    result["cod_amount"] = num_m.group(1)
                    result["shipment_value"] = num_m.group(1)
                matched = True
                break
        if matched: continue

        # Check Items count
        for pat in ITEMS_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                val = m.group(1).strip()
                num_m = re.search(r'(\d+)', val)
                if num_m:
                    result["items_count"] = num_m.group(1)
                matched = True
                break
        if matched: continue

        # Check Description
        for pat in DESC_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                result["description"] = m.group(1).strip()
                matched = True
                break
        if matched: continue

        # Check Notes
        for pat in NOTES_PATTERNS:
            m = re.match(pat, line, re.IGNORECASE)
            if m:
                result["notes"] = m.group(1).strip()
                matched = True
                break
        if matched: continue

        unhandled_lines.append(line)

    # Filter out lines that are purely phones, prices, codes, or common greetings from unhandled_lines
    clean_unhandled = []
    for line in unhandled_lines:
        line_clean = line.strip()
        if not line_clean:
            continue
        # Skip pure numbers/phones
        if re.sub(r'[^\d]', '', line_clean) in all_phones:
            continue
        # Skip greetings
        if line_clean in ['طلب جديد', 'اوردر جديد', 'أوردر جديد', 'أوردر', 'طلب', 'السلام عليكم', 'مساء الخير', 'صباح الخير']:
            continue
        # Skip lines that are just price like '350 جنيه' or 'السعر 350'
        if re.match(r'^(?:سعر الاوردر|السعر|المبلغ|الحساب)?\s*\d+\s*(?:ج|جنيه|ج\.م|egp|le)?$', line_clean, re.IGNORECASE):
            continue
        # Skip lines that are just items count like 'عدد 1' or 'عدد القطع 1'
        if re.match(r'^(?:عدد|العدد|عدد القطع|الكمية)?\s*\d+\s*(?:قطع|قطعة|قطع)?$', line_clean, re.IGNORECASE):
            continue
        # Skip lines that are codes or references (e.g. كود: GC01, كود الطلب: 123, ORD-1, #123)
        if re.match(r'^(?:كود(?:\s+الطلب|\s+الشحن|\s+المنتج)?|مرجع|المرجع|رمز|Code|Ref|SKU)\s*[:=\-#]?\s*[\w\-]+$', line_clean, re.IGNORECASE):
            continue
        if re.match(r'^(?:ORD[-_]?\d+|[A-Za-z]{1,4}[-_]?\d{1,6}|#\w+)$', line_clean, re.IGNORECASE):
            continue
        clean_unhandled.append(line_clean)
    unhandled_lines = clean_unhandled

    # If City or Area still empty, search in address or full text
    if not result["city"] or not result["area"]:
        search_target = result["address"] + " " + " ".join(unhandled_lines)
        c, a = match_city_and_area(search_target)
        if not result["city"] and c:
            result["city"] = c
        if not result["area"] and a:
            result["area"] = a

    # If Name still empty, guess from first unhandled line that looks like a name
    if not result["name"] and unhandled_lines:
        for idx, line in enumerate(unhandled_lines):
            cleaned_line = re.sub(r'^(?:السلام عليكم|مساء الخير|صباح الخير|طلب جديد|اوردر جديد|مرحبا|أوردر)\s*', '', line).strip()
            if not cleaned_line:
                continue
            # Must not be a phone or number or address
            if not re.search(r'\d{4,}', cleaned_line) and len(cleaned_line.split()) in [2, 3, 4, 5]:
                if not any(k in cleaned_line for k in ['جنيه', 'ج.م', 'EGP', 'شارع', 'ش ', 'عمارة', 'طنطا', 'القاهرة']):
                    result["name"] = cleaned_line
                    unhandled_lines.pop(idx)
                    break

    # If Address still empty, use remaining unhandled lines that contain street keywords
    if not result["address"] and unhandled_lines:
        addr_candidates = []
        for line in list(unhandled_lines):
            if any(w in line for w in ['شارع', 'ش ', 'عمارة', 'عماره', 'الدور', 'شقة', 'شقه', 'بجوار', 'امام', 'أمام', 'خلف', 'ميدان', 'طريق', 'قرية', 'مركز']):
                addr_candidates.append(line)
                unhandled_lines.remove(line)
        if addr_candidates:
            result["address"] = " - ".join(addr_candidates)
        elif unhandled_lines:
            # Use longest remaining line as address
            longest = max(unhandled_lines, key=len)
            if len(longest) > 10 and not result["address"]:
                result["address"] = longest
                unhandled_lines.remove(longest)

    # Check for price in remaining lines if not yet found
    if not result["cod_amount"]:
        price_m = re.search(r'(\d+)\s*(?:ج|جنيه|ج\.م|EGP|LE|LE\.)', text, re.IGNORECASE)
        if price_m:
            result["cod_amount"] = price_m.group(1)
            result["shipment_value"] = price_m.group(1)

    # If Description still empty, combine any remaining lines
    if not result["description"] and unhandled_lines:
        result["description"] = " - ".join(unhandled_lines)

    # Ensure Description has ONLY products and clean of any codes
    result["description"] = clean_product_text(result["description"])

    # Ensure Notes do not contain codes
    if result["notes"]:
        result["notes"] = clean_product_text(result["notes"])

    # Fallback shipment_value to cod_amount
    if not result["shipment_value"] and result["cod_amount"]:
        result["shipment_value"] = result["cod_amount"]

    # Validate phone
    if result["phone"]:
        is_valid, norm_phone, err = validate_egyptian_phone(result["phone"])
        if is_valid:
            result["phone"] = norm_phone
        else:
            result["errors"].append(err)
    else:
        result["errors"].append("لم يتم العثور على رقم هاتف في الرسالة!")

    # Validate City
    if not result["city"]:
        result["warnings"].append("لم يتم تحديد المحافظة (المدينة) بدقة - يرجى اختيارها من القائمة.")
    elif result["city"] not in BOSTA_CITIES:
        # Check alias
        alias_city = CITY_ALIASES.get(result["city"].lower())
        if alias_city in BOSTA_CITIES:
            result["city"] = alias_city
        else:
            result["warnings"].append(f"المحافظة '{result['city']}' غير مسجلة رسمياً في بوسطة - اختر من القائمة المتاحة.")

    # Validate Name
    if not result["name"]:
        result["errors"].append("اسم العميل مفقود! يرجى إدخال اسم العميل.")

    # Validate Address
    if not result["address"]:
        result["warnings"].append("العنوان غير محدد بالتفصيل.")

    return result


def parse_order_with_gemini(text: str, api_key: str) -> Dict[str, Any]:
    """
    Parses unstructured order text using Google Gemini AI.
    """
    try:
        from google import genai
        from google.genai import types

        client = genai.Client(api_key=api_key)
        
        cities_str = ", ".join(BOSTA_CITIES)
        
        prompt = f"""أنت خبير في استخراج بيانات طلبات التجارة الإلكترونية المصرية لمنصة الشحن 'بوسطة' (Bosta).
استخرج البيانات من الرسالة التالية بصيغة JSON حصراً:
```text
{text}
```

قائمة المحافظات المعتمدة في بوسطة هي فقط:
[{cities_str}]

المطلوب إرجاع كائن JSON بالحقول التالية بدقة:
{{
  "name": "اسم العميل كامل",
  "phone": "رقم الهاتف الأساسي 11 رقم يبدأ بـ 010 أو 011 أو 012 أو 015",
  "second_phone": "رقم هاتف إضافي إن وجد أو فارغ",
  "city": "اسم المحافظة حصراً من قائمة بوسطة المعتمدة",
  "area": "اسم المنطقة أو الحي",
  "address": "العنوان التفصيلي واسم الشارع والعمارة",
  "landmark": "علامة مميزة إن وجدت",
  "cod_amount": "قيمة التحصيل النقدي كرقم فقط",
  "items_count": "عدد القطع كرقم فقط (الافتراضي 1)",
  "description": "وصف المنتجات والقطع فقط بدون أي كود أو SKU أو رمز أو كود شحن",
  "notes": "أي ملاحظات خاصة بالتوصيل فقط وممنوع وضع أي كود طلب أو كود شحن هنا"
}}
أرجع JSON فقط بدون أي نصوص إضافية أو شروحات."""

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json"
            )
        )
        
        parsed_json = json.loads(response.text)
        
        # Clean and validate
        if "phone" in parsed_json and parsed_json["phone"]:
            parsed_json["phone"] = clean_phone(parsed_json["phone"])
            if len(parsed_json["phone"]) == 10 and parsed_json["phone"].startswith(('10', '11', '12', '15')):
                parsed_json["phone"] = '0' + parsed_json["phone"]
                
        if "second_phone" in parsed_json and parsed_json["second_phone"]:
            parsed_json["second_phone"] = clean_phone(parsed_json["second_phone"])

        # Ensure description has ONLY products
        parsed_json["description"] = clean_product_text(parsed_json.get("description", ""))
        
        # Ensure notes do not have codes
        if parsed_json.get("notes"):
            parsed_json["notes"] = clean_product_text(parsed_json["notes"])

        parsed_json["shipment_value"] = parsed_json.get("cod_amount", "")
        parsed_json["errors"] = []
        parsed_json["warnings"] = []
        parsed_json["raw_text"] = text

        # Validate phone
        if parsed_json.get("phone"):
            is_valid, norm_phone, err = validate_egyptian_phone(parsed_json["phone"])
            if is_valid:
                parsed_json["phone"] = norm_phone
            else:
                parsed_json["errors"].append(err)
        else:
            parsed_json["errors"].append("لم يتم العثور على رقم هاتف صالح!")

        return parsed_json

    except Exception as e:
        # Fallback to local regex parser if AI fails or error occurs
        local_res = parse_order_text(text)
        local_res["warnings"].append(f"تعذر الاتصال بالذكاء الاصطناعي ({str(e)}) - تم استخدام المحلل المدمج فائق السرعة.")
        return local_res
