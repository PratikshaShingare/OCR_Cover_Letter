"""
Khanna Travels & Holidays — ICAO 9303 TD3 (Passport) Machine Readable Zone (MRZ) Parser
Implements full check-digit validation, optical error recovery, and structured extraction.
TD3 standard format: 2 lines of 44 characters.
Line 1: P<UTOERIKSSON<<ANNA<MARIA<<<<<<<<<<<<<<<<<<<
Line 2: L898902C36UTO7408122F1204159ZE184226B<<<<<10
"""

import re
from typing import Dict, Any, Optional, Tuple
from datetime import datetime


def mrz_char_value(c: str) -> int:
    """Returns numerical value of character for ICAO Doc 9303 check digit calculation."""
    if '0' <= c <= '9':
        return ord(c) - ord('0')
    if 'A' <= c <= 'Z':
        return ord(c) - ord('A') + 10
    return 0  # '<' and fillers


def compute_check_digit(data_str: str) -> str:
    """Calculates ICAO 9303 check digit using weights 7, 3, 1 repeating modulo 10."""
    weights = [7, 3, 1]
    total = sum(mrz_char_value(c) * weights[i % 3] for i, c in enumerate(data_str))
    return str(total % 10)


def verify_check_digit(data_str: str, expected_digit: str) -> bool:
    """Returns True if check digit matches expected value."""
    if not expected_digit or not expected_digit.isdigit():
        return False
    return compute_check_digit(data_str) == expected_digit


def recover_field_with_check_digit(raw_field: str, expected_digit: str, is_numeric: bool = False) -> Tuple[str, bool]:
    """
    Attempts optical character confusion recovery using the mathematical check digit.
    Substitutions tested: O <-> 0, I <-> 1, Z <-> 2, S <-> 5, B <-> 8.
    """
    if verify_check_digit(raw_field, expected_digit):
        return raw_field, True

    confusions = [
        ('O', '0'), ('0', 'O'),
        ('I', '1'), ('1', 'I'),
        ('Z', '2'), ('2', 'Z'),
        ('S', '5'), ('5', 'S'),
        ('B', '8'), ('8', 'B'),
        ('G', '6'), ('6', 'G')
    ]

    chars = list(raw_field)
    for i, c in enumerate(chars):
        for orig, sub in confusions:
            if c == orig:
                chars[i] = sub
                test_str = "".join(chars)
                if verify_check_digit(test_str, expected_digit):
                    return test_str, True
                chars[i] = c  # revert

    return raw_field, False


def parse_mrz_date(date_str: str, is_expiry: bool = False) -> str:
    """Converts YYMMDD to standard YYYY-MM-DD format."""
    if not date_str or len(date_str) != 6 or not date_str.isdigit():
        return ""
    yy = int(date_str[0:2])
    mm = int(date_str[2:4])
    dd = int(date_str[4:6])

    # Validate basic date ranges
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return ""

    current_year = datetime.now().year % 100
    if is_expiry:
        year = 2000 + yy if yy <= (current_year + 35) else 1900 + yy
    else:
        year = 2000 + yy if yy <= current_year else 1900 + yy

    return f"{year:04d}-{mm:02d}-{dd:02d}"


def normalize_mrz_line(line: str) -> str:
    """Normalizes an OCR MRZ line by stripping whitespace and standardizing chevrons."""
    line = line.strip().upper().replace(" ", "")
    # Common OCR substitutions for chevron '<'
    for ch in ['«', '‹', '(', ')', '{', '}', '[', ']', '_']:
        line = line.replace(ch, '<')
    return line


def parse_mrz_lines(line1: str, line2: str) -> Optional[Dict[str, Any]]:
    """
    Parses and mathematically validates two lines of TD3 MRZ.
    Extracts all standard ICAO Doc 9303 fields with confidence scoring.
    """
    l1 = normalize_mrz_line(line1)
    l2 = normalize_mrz_line(line2)

    if len(l1) < 35 or len(l2) < 35:
        return None

    # Pad or trim to exactly 44 characters
    l1 = (l1 + "<" * 44)[:44]
    l2 = (l2 + "<" * 44)[:44]

    doc_type = l1[0:2].replace("<", "")
    issuing_country = l1[2:5].replace("<", "")

    # Line 1: Names start at position 5
    name_part = l1[5:]
    name_segments = name_part.split("<<")
    surname = name_segments[0].replace("<", " ").strip() if len(name_segments) > 0 else ""
    given_names = name_segments[1].replace("<", " ").strip() if len(name_segments) > 1 else ""
    full_name = f"{given_names} {surname}".strip() if surname else given_names

    # Clean any stray numbers from name fields (OCR artifact)
    surname = re.sub(r'[0-9]', '', surname).strip()
    given_names = re.sub(r'[0-9]', '', given_names).strip()
    full_name = f"{given_names} {surname}".strip() if surname else given_names

    # Line 2: Passport Number (0..8) and Check Digit (9)
    raw_passport = l2[0:9]
    pass_check = l2[9]
    passport_num, pass_valid = recover_field_with_check_digit(raw_passport, pass_check)
    passport_clean = passport_num.replace("<", "").strip()

    # Nationality: positions 10..12
    nationality = l2[10:13].replace("<", "")
    nat_clean = "Indian" if nationality in ["IND", "INDIA"] else (nationality or "Indian")

    # Date of Birth: positions 13..18 and Check Digit (19)
    raw_dob = l2[13:19]
    dob_check = l2[19]
    dob_fixed, dob_valid = recover_field_with_check_digit(raw_dob, dob_check, is_numeric=True)
    dob = parse_mrz_date(dob_fixed, is_expiry=False)

    # Sex: position 20
    sex_char = l2[20]
    if sex_char == "M":
        gender = "Male"
    elif sex_char == "F":
        gender = "Female"
    else:
        gender = "Other"

    # Expiry Date: positions 21..26 and Check Digit (27)
    raw_exp = l2[21:27]
    exp_check = l2[27]
    exp_fixed, exp_valid = recover_field_with_check_digit(raw_exp, exp_check, is_numeric=True)
    expiry_date = parse_mrz_date(exp_fixed, is_expiry=True)

    # Composite Check Digit: position 43
    # Check string: pos 0..9 (passport+chk) + pos 13..19 (dob+chk) + pos 21..42 (exp+chk+optional)
    composite_data = l2[0:10] + l2[13:20] + l2[21:43]
    composite_check = l2[43]
    composite_valid = verify_check_digit(composite_data, composite_check)

    # Calculate confidence based on check-digit validations
    checks_passed = sum([pass_valid, dob_valid, exp_valid, composite_valid])
    if checks_passed >= 3:
        confidence = 0.99
    elif checks_passed >= 1 or pass_valid:
        confidence = 0.92
    else:
        confidence = 0.80

    return {
        "passportType": doc_type or "P",
        "issuingCountry": "India" if issuing_country in ["IND", "INDIA"] else (issuing_country or "India"),
        "passportNumber": passport_clean,
        "surname": surname,
        "givenName": given_names,
        "fullName": full_name,
        "nationality": nat_clean,
        "dob": dob,
        "gender": gender,
        "expiryDate": expiry_date,
        "confidence": confidence,
        "checkDigits": {
            "passport": pass_valid,
            "dob": dob_valid,
            "expiry": exp_valid,
            "composite": composite_valid
        }
    }


def find_mrz_in_text(text: str) -> Optional[Dict[str, Any]]:
    """
    Scans OCR text for 2 lines matching TD3 MRZ pattern.
    Handles spaced-out characters and non-consecutive OCR text blocks.
    """
    raw_lines = [normalize_mrz_line(l) for l in text.splitlines() if l.strip()]
    candidates = [l for l in raw_lines if len(l) >= 35 and "<" in l]

    # 1. Consecutive candidates check
    for i in range(len(candidates) - 1):
        l1 = candidates[i]
        l2 = candidates[i + 1]
        if (l1.startswith("P<") or l1.startswith("P")) and ("<<" in l1 or len(l1) >= 40):
            res = parse_mrz_lines(l1, l2)
            if res and res.get("passportNumber"):
                return res

    # 2. Non-consecutive check: match line 1 starting with P< and line 2 starting with passport characters
    p_lines = [l for l in candidates if l.startswith("P<") or (l.startswith("P") and "<<" in l)]
    num_lines = [l for l in candidates if re.match(r'^[A-Z0-9]{7,9}[0-9<]', l)]

    if p_lines and num_lines:
        for p_l in p_lines:
            for n_l in num_lines:
                res = parse_mrz_lines(p_l, n_l)
                if res and res.get("passportNumber") and len(res["passportNumber"]) >= 7:
                    return res

    return None
