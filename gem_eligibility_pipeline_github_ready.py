#!/usr/bin/env python3
"""
gem_eligibility_pipeline.py
===========================
Phase-1/2 "Eligibility Extraction" pipeline for GeM tender intake.

This is built directly on top of the team's Google-Sheet-native pipeline
(reads the daily tab via HTTP CSV export or gspread, downloads each Bid
Document PDF, extracts + downloads nested child documents while filtering
out mailto: links, and can write results straight back into the Google
Sheet) -- with the eligibility engine upgraded to be structured and
requirement-level instead of a single free-text blob:

  - Every bid now gets a list of discrete requirement checks (turnover,
    product/OEM-specific turnover, experience, technical match, net worth,
    order count, order value, service activities, mandatory documents, ATC
    conditions) -- each with a Passed / Failed / Partially matched /
    Not found / Needs review status, an evidence snippet + source document,
    and a "mandatory" flag.
  - Product/OEM-specific turnover is evaluated separately from company-wide
    turnover, and can independently disqualify a bid.
  - A configurable scoring model (procmart_config.json) turns those checks
    into an eligibility score + AI confidence + Qualified / Segment One /
    Segment Two / Disqualified / Pending Human Review classification --
    with certain requirements ("blocking rules") able to force Disqualified
    regardless of the numeric score.
  - Every linked child document (ATC, technical spec, BOQ, SLA, GAR, and
    now Excel/Word attachments too) is downloaded, parsed, and reported on
    individually: processed / failed to download / failed to parse /
    unsupported type, plus warnings for low text density (possible scans)
    and regional-language (Hindi/Gujarati/Tamil) content.

Output:
  - The ORIGINAL Google Sheet columns are still produced and (optionally)
    written back for backward compatibility:
        BidDoc_eligibility, ChildDoc_eligibility, Procmart_eligibility,
        technical_spec_links
    plus new summary columns: Eligibility_Status, Eligibility_Score,
    AI_Confidence, Human_Review_Required, Eligibility_JSON (compact).
  - A rich `bids_<date>.json` file (one structured record per bid) is always
    written locally -- this is what the eligibility dashboard (index.html)
    consumes for the full requirement-by-requirement drill-down view.
  - A flattened `eligibility_output.csv` audit file, as before.
"""

import argparse
import csv
import io
import json
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Tuple, Any

import pdfplumber
import requests
from pypdf import PdfReader

# Google Sheets API client
try:
    import gspread
except ImportError:
    gspread = None

# Optional parsers for non-PDF child documents
try:
    import openpyxl
except Exception:
    openpyxl = None

try:
    import docx  # python-docx
except Exception:
    docx = None


# --------------------------------------------------------------------------
# Text cleaning helpers
# --------------------------------------------------------------------------

CID_RE = re.compile(r"\(cid:\d+\)")
DEVANAGARI_RE = re.compile(r"[\u0900-\u097F]+")
GUJARATI_RE = re.compile(r"[\u0A80-\u0AFF]+")
TAMIL_RE = re.compile(r"[\u0B80-\u0BFF]+")
WS_RE = re.compile(r"\s+")

REGIONAL_SCRIPTS = {
    "Hindi/Devanagari": DEVANAGARI_RE,
    "Gujarati": GUJARATI_RE,
    "Tamil": TAMIL_RE,
}


def clean_text(s: Optional[str]) -> str:
    """Strip PDF (cid:n) artifacts and Indic-script glyphs, collapse whitespace."""
    if not s:
        return ""
    s = CID_RE.sub("", s)
    for pat in REGIONAL_SCRIPTS.values():
        s = pat.sub("", s)
    s = WS_RE.sub(" ", s).strip()
    return s


def summarize_snippet(text: str, max_words: int = 15) -> str:
    """Summarizes text into a clean, concise single-line snippet."""
    if not text:
        return ""
    words = clean_text(text).split()
    if len(words) <= max_words:
        return " ".join(words)
    return " ".join(words[:max_words]) + "..."


def detect_regional_scripts(raw_text: str) -> List[str]:
    found = []
    if not raw_text:
        return found
    for name, pat in REGIONAL_SCRIPTS.items():
        if pat.search(raw_text):
            found.append(name)
    return found


# --------------------------------------------------------------------------
# PDF -> key/value table + full text extraction (in-memory stream)
# --------------------------------------------------------------------------

def extract_kv_rows_from_stream(pdf_file_obj: io.BytesIO) -> Dict[str, str]:
    """Extract 2-column key-value rows from an in-memory PDF file stream."""
    rows: Dict[str, str] = {}
    pdf_file_obj.seek(0)
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                for row in table:
                    if not row or len(row) < 2:
                        continue
                    label = clean_text(row[0])
                    value = clean_text(row[1]) if row[1] else ""
                    if label and len(label) < 400:
                        rows[label] = value
    return rows


def extract_full_text_from_stream(pdf_file_obj: io.BytesIO) -> Tuple[str, int, int]:
    """Returns (clean_full_text, raw_char_count, page_count) -- char/page counts
    are used for scanned-document / low-text-quality detection."""
    pdf_file_obj.seek(0)
    raw_parts = []
    with pdfplumber.open(pdf_file_obj) as pdf:
        page_count = len(pdf.pages)
        for p in pdf.pages:
            raw_parts.append(p.extract_text() or "")
    raw_text = "\n".join(raw_parts)
    return clean_text(raw_text), len(raw_text), page_count


def find_field(rows: Dict[str, str], *needles: str) -> Optional[str]:
    """Return matching row value where all search terms are present in the label."""
    for label, value in rows.items():
        low = label.lower()
        if all(n.lower() in low for n in needles):
            return value
    return None


def extract_bid_number(full_text: str, fallback: str) -> str:
    m = re.search(r"GEM/\d{4}/[A-Z]/\d+", full_text)
    return m.group(0) if m else fallback


def grab_short(full_text: str, pattern: str, max_words: int = 15) -> Optional[str]:
    """Find regex match and return a short concise snippet."""
    m = re.search(pattern, full_text, re.I)
    if not m:
        return None
    start = m.start()
    end = min(len(full_text), m.end() + 220)
    snippet = clean_text(full_text[start:end])
    return summarize_snippet(snippet, max_words)


def grab_number(text: str, pattern: str) -> Optional[float]:
    m = re.search(pattern, text, re.I)
    if not m:
        return None
    tail = text[m.end(): m.end() + 60]
    num_m = re.search(r"[\d,]+(?:\.\d+)?", tail)
    if not num_m:
        return None
    try:
        return float(num_m.group(0).replace(",", ""))
    except ValueError:
        return None


def parse_lakh_or_crore(value_str: Optional[str]) -> Optional[float]:
    """Convert strings like '10000 Lakh (s)' / '2 Crore' / '48000000' to INR."""
    if not value_str:
        return None
    s = value_str.replace(",", "")
    m = re.search(r"([\d.]+)\s*(lakh|lac|crore|cr)?", s, re.I)
    if not m:
        return None
    try:
        num = float(m.group(1))
    except ValueError:
        return None
    unit = (m.group(2) or "").lower()
    if unit.startswith("lakh") or unit.startswith("lac"):
        return num * 100000
    if unit.startswith("cr"):
        return num * 10000000
    return num


# --------------------------------------------------------------------------
# Expanded Link & Pattern Extraction (with mailto filtering)
# --------------------------------------------------------------------------

DOC_CATEGORY_PATTERNS = [
    ("technical_specification", re.compile(r"technical_specification|speccp|/spec", re.I)),
    ("boq_document", re.compile(r"BoqDocument", re.I)),
    ("boq_line_items_csv", re.compile(r"BoqLineItemsDocument", re.I)),
    ("atc_pqc_document", re.compile(r"/ATC|ATC_|ATC1_|ATCCP|pqc|prequal", re.I)),
    ("buyer_additional_doc", re.compile(r"downloadBuyerDoc|GWrfp|Buyer", re.I)),
    ("sla_bid_document", re.compile(r"bidsla", re.I)),
    ("gar_report", re.compile(r"GAR|gem_availability", re.I)),
    ("general_terms_conditions", re.compile(r"gtc/pdfByDate", re.I)),
    ("bid_summary_omp", re.compile(r"downloadOmppdfile", re.I)),
]

EMBEDDED_LINK_LABEL_RE = re.compile(
    r"([A-Za-z0-9\s/(),.\-\&]+?)\s*[:\-]?\s*(\d{6,}\.(?:pdf|xlsx|docx|csv))",
    re.I,
)


def categorize_link(url: str) -> str:
    for label, pattern in DOC_CATEGORY_PATTERNS:
        if pattern.search(url):
            return label
    return "other"


def extract_doc_links_from_stream(pdf_file_obj: io.BytesIO) -> List[Dict[str, str]]:
    """Extract external document links directly from the PDF stream annotations."""
    seen = set()
    links: List[Dict[str, str]] = []
    pdf_file_obj.seek(0)
    reader = PdfReader(pdf_file_obj)
    for page in reader.pages:
        annots = page.get("/Annots")
        if not annots:
            continue
        for a in annots:
            obj = a.get_object()
            if obj.get("/Subtype") != "/Link":
                continue
            action = obj.get("/A")
            if not action:
                continue
            uri = action.get("/URI")
            if not uri or uri in seen:
                continue
            seen.add(uri)
            links.append({"url": uri, "category": categorize_link(uri)})
    return links


def extract_all_hyperlinks_formatted(links: List[Dict[str, str]], full_text: str) -> Tuple[str, List[Tuple[str, str]]]:
    """
    Extracts all attached links, filtering out mailto: links, associates them with labels,
    and formats each link as 'Label : URL' listed one under the other.
    Returns (formatted_string, [(label, url), ...]).
    """
    formatted_entries = []
    link_pairs = []
    seen_urls = set()

    label_map = {}
    for match in EMBEDDED_LINK_LABEL_RE.finditer(full_text):
        label = clean_text(match.group(1))
        doc_file = match.group(2)
        doc_id = doc_file.split(".")[0]
        if label and doc_id:
            label_map[doc_id] = (label, doc_file)

    for l in links:
        url = l["url"]

        # --- FILTER OUT MAILTO & EMPTY LINKS ---
        if not url or url.lower().startswith("mailto:") or url in seen_urls:
            continue

        seen_urls.add(url)

        matched_label = None
        for doc_id, (lbl, doc_file) in label_map.items():
            if doc_id in url:
                matched_label = f"{lbl} ({doc_file})"
                break

        if not matched_label:
            matched_label = f"Document ({l['category'].replace('_', ' ').title()})"

        formatted_entries.append(f"{matched_label} : {url}")
        link_pairs.append((matched_label, url))

    formatted_output = ";\n".join(formatted_entries)
    return formatted_output, link_pairs


# --------------------------------------------------------------------------
# Child-document fetch + parse (PDF / XLSX / DOCX / CSV)
# --------------------------------------------------------------------------

@dataclass
class ChildDocResult:
    name: str
    url: str
    doc_type: str
    status: str  # processed | failed_download | failed_parse | unsupported_type
    text: str = ""
    char_count: int = 0
    page_count: int = 0
    warnings: List[str] = field(default_factory=list)
    requirements_found: List[str] = field(default_factory=list)


def guess_doc_type(url: str) -> str:
    low = url.lower().split("?")[0]
    if low.endswith(".pdf"):
        return "PDF"
    if low.endswith(".xlsx") or low.endswith(".xls"):
        return "Excel"
    if low.endswith(".docx") or low.endswith(".doc"):
        return "Word"
    if low.endswith(".csv"):
        return "CSV"
    return "Unknown"


def fetch_and_parse_child_doc(session: requests.Session, label: str, url: str, timeout: int = 20) -> ChildDocResult:
    doc_type = guess_doc_type(url)
    result = ChildDocResult(name=label, url=url, doc_type=doc_type, status="failed_download")

    try:
        resp = session.get(url, timeout=timeout)
    except Exception as e:
        result.warnings.append(f"Download error: {e}")
        return result

    if resp.status_code != 200 or not resp.content:
        result.warnings.append(f"HTTP {resp.status_code}")
        return result

    content = resp.content

    try:
        if doc_type == "PDF":
            stream = io.BytesIO(content)
            clean_full, raw_chars, pages = extract_full_text_from_stream(stream)
            result.text = clean_full
            result.char_count = raw_chars
            result.page_count = pages
            result.status = "processed"
            if pages > 0 and (raw_chars / max(pages, 1)) < 40:
                result.warnings.append("Low text extraction / possibly scanned image (needs OCR review)")

        elif doc_type == "Excel" and openpyxl is not None:
            wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True, read_only=True)
            parts = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    parts.append(" ".join(str(c) for c in row if c is not None))
            result.text = clean_text(" \n ".join(parts))
            result.char_count = len(result.text)
            result.page_count = len(wb.worksheets)
            result.status = "processed"

        elif doc_type == "Word" and docx is not None:
            d = docx.Document(io.BytesIO(content))
            result.text = clean_text("\n".join(p.text for p in d.paragraphs))
            result.char_count = len(result.text)
            result.page_count = 1
            result.status = "processed"

        elif doc_type == "CSV":
            result.text = clean_text(content.decode("utf-8", errors="ignore"))
            result.char_count = len(result.text)
            result.status = "processed"

        else:
            result.status = "unsupported_type"
            result.warnings.append(f"No parser available for type: {doc_type}")
            return result

    except Exception as e:
        result.status = "failed_parse"
        result.warnings.append(f"Parse error: {e}")
        return result

    for reqlabel, pat in [
        ("Turnover clause", r"turnover"),
        ("Experience / PQC clause", r"experience criteria|pre[-\s]?qualif|past experience"),
        ("Net worth clause", r"net\s*worth"),
        ("Order count / value clause", r"single order|two orders|three orders|number of orders"),
        ("Payment terms", r"payment terms|milestone"),
        ("Penalty / LD clause", r"liquidated damages|penalty"),
        ("Mandatory certification", r"iso\s*9001|udyam|dpiit|bis licen[cs]e|type test"),
        ("Scope of work / service activities", r"scope of work|installation.{0,20}commissioning|operation\s*&\s*maintenance"),
    ]:
        if re.search(pat, result.text, re.I):
            result.requirements_found.append(reqlabel)

    return result


# --------------------------------------------------------------------------
# Requirement model (structured, requirement-level eligibility)
# --------------------------------------------------------------------------

STATUS_PASSED = "Passed"
STATUS_FAILED = "Failed"
STATUS_PARTIAL = "Partially matched"
STATUS_NOT_FOUND = "Not found"
STATUS_REVIEW = "Needs review"

COLOR_FOR_STATUS = {
    STATUS_PASSED: "green",
    STATUS_FAILED: "red",
    STATUS_PARTIAL: "amber",
    STATUS_REVIEW: "amber",
    STATUS_NOT_FOUND: "grey",
}


@dataclass
class Requirement:
    id: str
    name: str
    category: str
    required_value: str
    status: str
    mandatory: bool
    evidence_doc: str
    evidence_snippet: str
    notes: str = ""

    def to_dict(self):
        d = self.__dict__.copy()
        d["color"] = COLOR_FOR_STATUS.get(self.status, "grey")
        return d


REQ_WEIGHTS = {
    "turnover": 12, "product_turnover": 18, "experience": 15, "technical_specification": 15,
    "net_worth": 8, "order_count": 12, "order_value": 8, "service_activities": 4,
    "mandatory_documents": 5, "atc_conditions": 3,
}


# --------------------------------------------------------------------------
# Business-section / govt-classification helpers
# --------------------------------------------------------------------------

def classify_business_section(item_category: str, full_text: str, config: dict) -> str:
    text = f"{item_category} {full_text[:3000]}".lower()
    sections = config.get("business_sections", {})
    best_section, best_hits = "Unclassified", 0
    for section, keywords in sections.items():
        hits = sum(1 for kw in keywords if kw in text)
        if hits > best_hits:
            best_section, best_hits = section, hits
    return best_section


def classify_govt(ministry: str, organisation: str, department: str, config: dict) -> str:
    hints = config.get("govt_classification_hints", {})
    blob = f"{ministry} {organisation} {department}".lower()
    for kw in hints.get("state_keywords", []):
        if kw in blob:
            return "State"
    for kw in hints.get("psu_keywords", []):
        if kw in blob:
            return "PSU"
    return "Central"


def detect_qcbs(full_text: str) -> bool:
    return bool(re.search(r"\bQCBS\b|quality\s*and\s*cost\s*based\s*selection", full_text, re.I))


# --------------------------------------------------------------------------
# Requirement-level evaluation
# --------------------------------------------------------------------------

def evaluate_requirements(rows: Dict[str, str], full_text: str, child_docs: List[ChildDocResult],
                           section: str, config: dict) -> List[Requirement]:
    company = config.get("company_profile", {})
    policy = config.get("policy", {})
    child_text_all = "\n".join(c.text for c in child_docs if c.text)
    combined_text = full_text + "\n" + child_text_all

    reqs: List[Requirement] = []

    # ---- 1. Company Turnover ----
    turnover_field = find_field(rows, "minimum average annual turnover") or find_field(rows, "turnover of the bidder")
    turnover_required_inr = parse_lakh_or_crore(turnover_field)
    company_turnover = company.get("annual_turnover_inr", 0)
    if turnover_field is None:
        status = STATUS_REVIEW if policy.get("treat_missing_turnover_field_as") == "NEEDS_REVIEW" else STATUS_NOT_FOUND
        reqs.append(Requirement("turnover", "Company Turnover", "Financial", "Not specified", status, False,
                                 "Main Bid Document", "No turnover clause detected"))
    else:
        passed = turnover_required_inr is not None and company_turnover >= turnover_required_inr
        reqs.append(Requirement(
            "turnover", "Company Turnover", "Financial",
            f"INR {turnover_required_inr:,.0f}" if turnover_required_inr else turnover_field,
            STATUS_PASSED if passed else STATUS_FAILED, False,
            "Main Bid Document", summarize_snippet(turnover_field, 15),
            notes=f"ProcMart turnover on record: INR {company_turnover:,.0f}"
        ))

    # ---- 2. Product / OEM-specific turnover (kept separate from company turnover) ----
    oem_turnover_field = find_field(rows, "oem average turnover") or find_field(rows, "oem turn over")
    oem_clause = grab_short(combined_text, r"OEM Turn Over Criteria", 30)
    product_specific_required = oem_turnover_field is not None or bool(oem_clause)
    if product_specific_required:
        req_inr = parse_lakh_or_crore(oem_turnover_field) if oem_turnover_field else None
        section_turnover_map = company.get("product_specific_turnover_inr", {})
        company_product_turnover = section_turnover_map.get(section, section_turnover_map.get("_default", 0))
        if req_inr is None:
            reqs.append(Requirement("product_turnover", "Product/OEM-Specific Turnover", "Financial",
                                     oem_turnover_field or (oem_clause or "Referenced in ATC"),
                                     STATUS_REVIEW, True, "Main Bid Document / ATC",
                                     oem_clause or "OEM turnover clause present but value unclear"))
        else:
            passed = company_product_turnover >= req_inr
            reqs.append(Requirement(
                "product_turnover", "Product/OEM-Specific Turnover", "Financial",
                f"INR {req_inr:,.0f}",
                STATUS_PASSED if passed else STATUS_FAILED, True,
                "Main Bid Document",
                oem_turnover_field or "",
                notes=(f"ProcMart {section} product turnover on record: INR {company_product_turnover:,.0f}. "
                       f"Company-wide turnover alone is NOT sufficient for this bid.")
            ))
    else:
        reqs.append(Requirement("product_turnover", "Product/OEM-Specific Turnover", "Financial",
                                 "Not required", STATUS_NOT_FOUND, False,
                                 "Main Bid Document", "No product-specific turnover clause found"))

    # ---- 3. Experience ----
    exp_field = find_field(rows, "years of past experience required")
    exp_clause = grab_short(combined_text, r"Experience Criteria\s*:", 25)
    years_required = None
    if exp_field:
        m = re.search(r"(\d+(?:\.\d+)?)", exp_field)
        years_required = float(m.group(1)) if m else None
    if exp_field is None and not exp_clause:
        status = STATUS_REVIEW if policy.get("treat_missing_experience_field_as") == "NEEDS_REVIEW" else STATUS_NOT_FOUND
        reqs.append(Requirement("experience", "Years of Past Experience", "Experience", "Not specified",
                                 status, False, "Main Bid Document", "No experience clause detected"))
    else:
        company_years = company.get("years_in_operation", 0)
        if years_required is not None:
            passed = company_years >= years_required
            reqs.append(Requirement(
                "experience", "Years of Past Experience", "Experience",
                f"{years_required:g} year(s)",
                STATUS_PASSED if passed else STATUS_FAILED, True,
                "Main Bid Document", exp_field or exp_clause or "",
                notes=f"ProcMart operating history: {company_years} year(s)"
            ))
        else:
            reqs.append(Requirement("experience", "Years of Past Experience", "Experience",
                                     exp_field or exp_clause, STATUS_REVIEW, True,
                                     "Main Bid Document", exp_clause or "", notes="Value could not be parsed automatically"))

    # ---- 4. Technical capability / product-equivalence ----
    tech_clause = grab_short(combined_text, r"Technical Specification|As per GeM Category Specification", 20)
    section_keywords = config.get("business_sections", {}).get(section, [])
    matched_kw = [kw for kw in section_keywords if kw in combined_text.lower()]
    if section == "Unclassified" or not matched_kw:
        reqs.append(Requirement("technical_specification", "Technical Capability / Product Match", "Technical",
                                 tech_clause or "See technical spec table", STATUS_REVIEW, True,
                                 "Main Bid Document / Technical Spec", tech_clause or "",
                                 notes="Product/service equivalence could not be confidently matched to ProcMart catalog"))
    else:
        reqs.append(Requirement("technical_specification", "Technical Capability / Product Match", "Technical",
                                 tech_clause or ", ".join(matched_kw), STATUS_PASSED, True,
                                 "Main Bid Document / Technical Spec", ", ".join(matched_kw),
                                 notes=f"Matched section keywords: {', '.join(matched_kw[:6])}"))

    # ---- 5. Net worth ----
    nw_clause = grab_short(combined_text, r"NET\s*WORTH", 20)
    if nw_clause:
        positive_required = bool(re.search(r"positive", nw_clause, re.I))
        company_nw = company.get("net_worth_inr", 0)
        passed = (company_nw > 0) if positive_required else True
        reqs.append(Requirement("net_worth", "Net Worth", "Financial", nw_clause,
                                 STATUS_PASSED if passed else STATUS_FAILED, True,
                                 "ATC / Buyer Added Terms", nw_clause,
                                 notes=f"ProcMart net worth on record: INR {company_nw:,.0f}"))
    else:
        reqs.append(Requirement("net_worth", "Net Worth", "Financial", "Not specified",
                                 STATUS_NOT_FOUND, False, "ATC", "No net worth clause found"))

    # ---- 6. Number of completed orders ----
    order_clause = grab_short(combined_text, r"single order of at least|two orders of at least|three orders of at least|number of orders", 35)
    if order_clause:
        company_orders = company.get("completed_orders_by_section", {}).get(section,
                          company.get("completed_orders_by_section", {}).get("_default", 0))
        passed = company_orders >= 3
        reqs.append(Requirement("order_count", "Number of Completed Orders", "Past Performance", order_clause,
                                 STATUS_PASSED if passed else STATUS_FAILED, True,
                                 "ATC / Buyer Added Terms", order_clause,
                                 notes=f"ProcMart completed orders on record ({section}): {company_orders}"))
    else:
        reqs.append(Requirement("order_count", "Number of Completed Orders", "Past Performance", "Not specified",
                                 STATUS_NOT_FOUND, False, "ATC", "No explicit order-count clause found"))

    # ---- 7. Order quantity / project value (past performance %) ----
    pp_field = find_field(rows, "past performance")
    pp_clause = pp_field or grab_short(combined_text, r"Past Performance\s*:", 20)
    if pp_clause:
        reqs.append(Requirement("order_value", "Order Quantity / Project Value (Past Performance)", "Past Performance",
                                 pp_clause, STATUS_REVIEW, False, "Main Bid Document / ATC", pp_clause,
                                 notes="Requires matching completed order value/quantity against bid quantity"))
    else:
        reqs.append(Requirement("order_value", "Order Quantity / Project Value (Past Performance)", "Past Performance",
                                 "Not specified", STATUS_NOT_FOUND, False, "Main Bid Document", ""))

    # ---- 8. Required service activities (supply/install/test/commission/AMC) ----
    activities = []
    for label, pat in [
        ("Supply", r"\bsupply\b"),
        ("Installation", r"\binstallation\b"),
        ("Testing", r"\btesting\b"),
        ("Commissioning", r"\bcommissioning\b"),
        ("Operation & Maintenance", r"operation\s*&?\s*maintenance|\bO&M\b"),
        ("After-sales / Warranty Service", r"after[-\s]sales|warranty service|amc\b"),
    ]:
        if re.search(pat, combined_text, re.I):
            activities.append(label)
    if activities:
        reqs.append(Requirement("service_activities", "Required Service Activities", "Scope",
                                 ", ".join(activities), STATUS_PASSED, False,
                                 "Main Bid Document / Scope of Work", ", ".join(activities)))
    else:
        reqs.append(Requirement("service_activities", "Required Service Activities", "Scope",
                                 "Supply only (default)", STATUS_NOT_FOUND, False,
                                 "Main Bid Document", ""))

    # ---- 9. Mandatory documents ----
    docs_field = find_field(rows, "document required from seller")
    if docs_field:
        doc_list = [d.strip() for d in re.split(r",", docs_field) if d.strip()]
        reqs.append(Requirement("mandatory_documents", "Required Supporting Documents", "Compliance",
                                 f"{len(doc_list)} document(s): " + summarize_snippet(docs_field, 20),
                                 STATUS_REVIEW, True, "Main Bid Document", docs_field,
                                 notes="Confirm each listed document is on file before submission"))
    else:
        reqs.append(Requirement("mandatory_documents", "Required Supporting Documents", "Compliance",
                                 "Not specified", STATUS_NOT_FOUND, False, "Main Bid Document", ""))

    # ---- 10. Additional ATC / hidden conditions ----
    atc_flags = []
    for label, pat in [
        ("EMD mandatory", r"without emd.{0,40}(rejected|summarily)"),
        ("MSE reserved", r"reserved for mse"),
        ("OEM authorization mandatory", r"oem authorization certificate|manufacturer authorization"),
        ("Local content / Make-in-India", r"class\s*1.{0,10}local supplier|local content"),
        ("Land-border restriction", r"land border with india"),
        ("Integrity Pact mandatory", r"integrity pact"),
        ("Sample / trial required", r"sample.{0,20}approv|advance sample"),
        ("Inspection required", r"inspection required"),
    ]:
        if re.search(pat, combined_text, re.I):
            atc_flags.append(label)
    reqs.append(Requirement("atc_conditions", "Additional ATC / Linked-Document Conditions", "Compliance",
                             ", ".join(atc_flags) if atc_flags else "None detected",
                             STATUS_REVIEW if atc_flags else STATUS_NOT_FOUND, False,
                             "ATC / Buyer Added Bid Specific Terms", ", ".join(atc_flags)))

    return reqs


# --------------------------------------------------------------------------
# Scoring, segmentation, mandatory-rule overrides
# --------------------------------------------------------------------------

def score_requirements(reqs: List[Requirement], config: dict) -> Dict[str, Any]:
    policy = config.get("policy", {})
    blocking_ids = set(policy.get("blocking_requirement_ids", []))

    total_weight, earned, contributing, reducing = 0, 0, [], []
    blocking_failed = []

    for r in reqs:
        w = REQ_WEIGHTS.get(r.id, 5)
        total_weight += w
        if r.status == STATUS_PASSED:
            earned += w
            contributing.append(f"{r.name} passed (+{w})")
        elif r.status == STATUS_PARTIAL:
            earned += w * 0.5
            contributing.append(f"{r.name} partially matched (+{w*0.5:.0f})")
            reducing.append(f"{r.name} only partially matched (-{w*0.5:.0f})")
        elif r.status == STATUS_FAILED:
            reducing.append(f"{r.name} failed (-{w})")
            if r.id in blocking_ids and r.mandatory:
                blocking_failed.append(r.name)
        elif r.status == STATUS_REVIEW:
            earned += w * 0.4
            reducing.append(f"{r.name} needs review (uncertain, +{w*0.4:.0f} only)")
        elif r.status == STATUS_NOT_FOUND and not r.mandatory:
            total_weight -= w * 0.5

    score = round(100 * earned / total_weight, 1) if total_weight > 0 else 0.0
    return {
        "eligibility_score": score,
        "contributing_factors": contributing,
        "reducing_factors": reducing,
        "blocking_conditions_failed": blocking_failed,
        "has_blocking_failure": len(blocking_failed) > 0,
    }


def compute_ai_confidence(rows: Dict[str, str], full_text: str, child_docs: List[ChildDocResult],
                           reqs: List[Requirement]) -> Tuple[float, List[str]]:
    warnings = []
    confidence = 90.0

    if len(rows) < 5:
        confidence -= 15
        warnings.append("Very few structured fields extracted from main bid document")
    if len(full_text) < 800:
        confidence -= 15
        warnings.append("Main bid document text extraction looks unusually short")

    for c in child_docs:
        if c.status != "processed":
            confidence -= 4
        for w in c.warnings:
            if "OCR" in w or "scanned" in w.lower() or "Low text" in w:
                confidence -= 8
                warnings.append(f"{c.name}: {w}")

    n_review = sum(1 for r in reqs if r.status == STATUS_REVIEW)
    n_notfound_mandatory = sum(1 for r in reqs if r.mandatory and r.status == STATUS_NOT_FOUND)
    confidence -= n_review * 3
    confidence -= n_notfound_mandatory * 4

    confidence = max(5.0, min(98.0, confidence))
    return round(confidence, 1), warnings


def classify_status(score_info: Dict[str, Any], ai_confidence: float, config: dict,
                     regional_langs: List[str], doc_quality_warnings: List[str], qcbs: bool) -> Tuple[str, str, bool]:
    """Returns (status, reason, human_review_required)."""
    thr = config.get("thresholds", {})
    policy = config.get("policy", {})

    if score_info["has_blocking_failure"]:
        reason = "Disqualified - blocking condition(s) failed: " + "; ".join(score_info["blocking_conditions_failed"])
        return "Disqualified", reason, False

    review_needed = False
    review_reasons = []

    if ai_confidence <= thr.get("review_confidence_max", 65) and ai_confidence >= thr.get("review_confidence_min", 0):
        review_needed = True
        review_reasons.append(f"AI confidence {ai_confidence}% is in the manual-review band")
    if policy.get("regional_language_review", True) and regional_langs:
        review_needed = True
        review_reasons.append(f"Regional language content detected ({', '.join(regional_langs)})")
    if policy.get("low_text_extraction_review", True) and doc_quality_warnings:
        review_needed = True
        review_reasons.append("Document quality warning(s) present")
    if qcbs:
        review_needed = True
        review_reasons.append("QCBS / special evaluation method - requires separate technical-scoring review")

    score = score_info["eligibility_score"]
    if score >= thr.get("qualified_min_score", 85):
        base_status = "Qualified"
    elif score >= thr.get("segment_one_min_score", 70):
        base_status = "Segment One - Review"
    elif score >= thr.get("segment_two_min_score", 50):
        base_status = "Segment Two - Low Match"
    else:
        base_status = "Disqualified"

    if review_needed and base_status != "Disqualified":
        status = "Pending Human Review"
        reason = f"Score {score}% ({base_status}); " + "; ".join(review_reasons)
        return status, reason, True

    reason = f"Score {score}% classified as {base_status} against configured thresholds"
    return base_status, reason, False


# --------------------------------------------------------------------------
# Legacy asterisk-format summaries (kept for backward-compatible Sheet columns)
# --------------------------------------------------------------------------

def legacy_biddoc_eligibility(rows: Dict[str, str], full_text: str) -> str:
    results: Dict[str, str] = {}

    cat = find_field(rows, "item category") or find_field(rows, "boq title") or find_field(rows, "primary product category")
    if cat:
        results["Product/Service Eligibility"] = summarize_snippet(cat, 12)

    turnover = find_field(rows, "minimum average annual turnover") or find_field(rows, "oem average turnover")
    if turnover:
        results["Turnover Eligibility"] = summarize_snippet(turnover, 10)

    exp = find_field(rows, "years of past experience required") or grab_short(full_text, r"Experience Criteria\s*:", 10)
    if exp:
        results["Experience Eligibility"] = exp

    pp = find_field(rows, "past performance") or grab_short(full_text, r"Past Performance\s*:", 10)
    if pp:
        results["Past Performance Eligibility"] = pp

    oem = grab_short(full_text, r"Manufacturer Authorization|must be the manufacturer|OEM Authorization", 10)
    val = find_field(rows, "document required from seller")
    if val and "OEM Authorization" in val:
        oem = "OEM Authorization Certificate Mandatory"
    if oem:
        results["OEM/Manufacturer Eligibility"] = oem

    certs = []
    for cert in ["UDYAM", "GSTIN", "DPIIT", "ISI Marked", "ISO"]:
        if re.search(r"\b" + cert + r"\b", full_text, re.I):
            certs.append(cert)
    if certs:
        results["Mandatory Certifications & Registrations"] = ", ".join(certs)

    pqc = grab_short(full_text, r"PRE[-\s]?QUALIFYING CONDITIONS|List of Required documents", 12)
    if pqc:
        results["Bid-Specific PQC/ATC Compliance"] = pqc

    emd = find_field(rows, "emd amount")
    if emd:
        results["EMD/Exemption Eligibility"] = f"INR {emd}"
    elif re.search(r"EMD EXEMPTION", full_text, re.I):
        results["EMD/Exemption Eligibility"] = "Exemption Applicable per GeM GTC"

    mii = find_field(rows, "mii purchase preference") or find_field(rows, "mse purchase preference")
    if mii:
        results["MII/MSE/Startup Conditions"] = f"Preference: {mii}"

    docs = find_field(rows, "document required from seller")
    if docs:
        results["Mandatory Documentation"] = summarize_snippet(docs, 12)

    lines = [f"*{k}: {v}" for k, v in results.items()]
    return "\n".join(lines)


def legacy_childdoc_eligibility(sub_docs: List[Tuple[str, str]]) -> str:
    if not sub_docs:
        return "*No nested child documents found or processed"

    results: Dict[str, str] = {}
    for label, text in sub_docs:
        if not text:
            continue
        turnover = grab_short(text, r"Turnover Criteria|Annual Turnover", 10)
        if turnover and "Turnover Eligibility" not in results:
            results["Turnover Eligibility"] = turnover
        pqc = grab_short(text, r"Pre-qualification|PQC|Eligibility Criteria", 12)
        if pqc and "PQC Extractions" not in results:
            results["PQC Extractions"] = pqc
        payment = grab_short(text, r"Payment Terms|Payment Condition|Milestone", 12)
        if payment and "Payment Terms Criteria" not in results:
            results["Payment Terms Criteria"] = payment
        scope = grab_short(text, r"Scope of Work|Technical Scope|Deliverables", 12)
        if scope and "Scope Requirements" not in results:
            results["Scope Requirements"] = scope
        sla = grab_short(text, r"Penalty|SLA|Liquidated Damages", 12)
        if sla and "SLA & Penalty Clauses" not in results:
            results["SLA & Penalty Clauses"] = sla
        cert = grab_short(text, r"Certifications Required|ISO|Quality Certificate", 10)
        if cert and "Additional Certifications" not in results:
            results["Additional Certifications"] = cert

    if not results:
        results["Child Doc Review"] = "No additional restrictive eligibility clauses identified in attached files."

    lines = [f"*{k}: {v}" for k, v in results.items()]
    return "\n".join(lines)


def legacy_procmart_summary(status: str, reqs: List[Requirement]) -> str:
    """Renders the structured requirement list back into the old asterisk format,
    so downstream sheet consumers relying on 'Procmart_eligibility' keep working."""
    lines = [f"*Overall ProcMart Qualification: {status.upper()}"]
    for r in reqs:
        lines.append(f"*{r.name}: {r.status.upper()} - {r.required_value}")
    return "\n".join(lines)


# --------------------------------------------------------------------------
# Dataclass & Processing Core
# --------------------------------------------------------------------------

@dataclass
class BidRecord:
    bid_number: str
    bid_doc_url: str = ""
    bid_doc_eligibility: str = ""
    child_doc_eligibility: str = ""
    procmart_eligibility: str = ""
    technical_spec_links: str = ""
    status: str = ""
    eligibility_score: float = 0.0
    ai_confidence: float = 0.0
    human_review_required: bool = False
    rich: Dict[str, Any] = field(default_factory=dict)


def parse_bid_pdf_stream(pdf_bytes: bytes, original_bid_no: str, doc_url: str, session: requests.Session,
                          config: dict, sheet_meta: Optional[Dict[str, str]] = None,
                          fetch_children: bool = True) -> BidRecord:
    sheet_meta = sheet_meta or {}
    raw_text_probe = pdf_bytes.decode("latin-1", errors="ignore")
    regional_langs = detect_regional_scripts(raw_text_probe)

    pdf_stream = io.BytesIO(pdf_bytes)
    rows = extract_kv_rows_from_stream(pdf_stream)
    full_text, raw_chars, page_count = extract_full_text_from_stream(pdf_stream)
    links = extract_doc_links_from_stream(pdf_stream)
    formatted_hyperlinks, link_pairs = extract_all_hyperlinks_formatted(links, full_text)

    doc_quality_warnings = []
    if page_count > 0 and (raw_chars / max(page_count, 1)) < 40:
        doc_quality_warnings.append("Main bid document has low text density (possible scan)")

    # ---- child documents (download + structured parse) ----
    child_results: List[ChildDocResult] = []
    legacy_sub_docs: List[Tuple[str, str]] = []
    if fetch_children:
        for label, sub_url in link_pairs:
            if sub_url.endswith(".pdf") or "download" in sub_url.lower() or sub_url.endswith((".xlsx", ".docx", ".csv")):
                cr = fetch_and_parse_child_doc(session, label, sub_url)
                child_results.append(cr)
                doc_quality_warnings.extend(cr.warnings)
                if cr.text:
                    legacy_sub_docs.append((label, cr.text))

    n_linked = len(link_pairs)
    n_processed = sum(1 for c in child_results if c.status == "processed")
    n_failed = len(child_results) - n_processed

    # ---- classification ----
    item_category = (find_field(rows, "item category") or find_field(rows, "boq title")
                      or sheet_meta.get("Title / Items", "") or "")
    ministry = find_field(rows, "ministry") or sheet_meta.get("Ministry", "") or ""
    organisation = find_field(rows, "organisation") or ""
    department = find_field(rows, "department") or sheet_meta.get("Department", "") or ""
    section = classify_business_section(item_category, full_text, config)
    govt_class = classify_govt(ministry, organisation, department, config)
    qcbs = detect_qcbs(full_text)

    bid_value = (parse_lakh_or_crore(find_field(rows, "estimated bid value")) or
                 parse_lakh_or_crore(find_field(rows, "estimated bid value in inr")))
    emd_amount = parse_lakh_or_crore(find_field(rows, "emd amount"))

    # ---- structured requirement-level evaluation ----
    reqs = evaluate_requirements(rows, full_text, child_results, section, config)
    score_info = score_requirements(reqs, config)
    ai_confidence, conf_warnings = compute_ai_confidence(rows, full_text, child_results, reqs)
    status, status_reason, needs_review = classify_status(
        score_info, ai_confidence, config, regional_langs, doc_quality_warnings, qcbs
    )

    bid_number = extract_bid_number(full_text, original_bid_no)

    # ---- legacy text summaries (backward compatible sheet columns) ----
    bid_doc_elig = legacy_biddoc_eligibility(rows, full_text)
    child_doc_elig = legacy_childdoc_eligibility(legacy_sub_docs)
    procmart_elig = legacy_procmart_summary(status, reqs)

    rich = {
        "bid_number": bid_number,
        "bid_doc_url": doc_url,
        "item_category": item_category,
        "ministry": ministry,
        "organisation": organisation,
        "department": department,
        "govt_classification": govt_class,
        "business_section": section,
        "bid_value_inr": bid_value,
        "emd_amount_inr": emd_amount,
        "evaluation_method": find_field(rows, "evaluation method") or "",
        "is_qcbs_or_special": qcbs,
        "end_date_raw": find_field(rows, "bid end date"),

        "status": status,
        "status_reason": status_reason,
        "eligibility_score": score_info["eligibility_score"],
        "ai_confidence": ai_confidence,
        "human_review_required": needs_review,
        "human_review_reasons": (conf_warnings + (
            [f"Regional language: {l}" for l in regional_langs] if regional_langs else []
        )),

        "scoring": {
            "eligibility_score": score_info["eligibility_score"],
            "ai_confidence": ai_confidence,
            "contributing_factors": score_info["contributing_factors"],
            "reducing_factors": score_info["reducing_factors"],
            "blocking_conditions_failed": score_info["blocking_conditions_failed"],
            "has_blocking_failure": score_info["has_blocking_failure"],
            "classification_basis": "hard_rule" if score_info["has_blocking_failure"] else "weighted_requirement_score",
        },

        "requirements": [r.to_dict() for r in reqs],

        "documents": {
            "main_document_quality_warnings": doc_quality_warnings,
            "linked_documents_found": n_linked,
            "linked_documents_processed": n_processed,
            "linked_documents_failed": n_failed,
            "regional_language_detected": regional_langs,
            "children": [
                {
                    "name": c.name, "url": c.url, "type": c.doc_type, "status": c.status,
                    "char_count": c.char_count, "page_count": c.page_count,
                    "warnings": c.warnings, "requirements_found": c.requirements_found,
                }
                for c in child_results
            ],
        },

        "technical_spec_links": formatted_hyperlinks,
        "human_review": {"resolved": False, "final_status": None, "reviewer_comment": "", "override_reason": ""},
    }

    return BidRecord(
        bid_number=bid_number,
        bid_doc_url=doc_url,
        bid_doc_eligibility=bid_doc_elig,
        child_doc_eligibility=child_doc_elig,
        procmart_eligibility=procmart_elig,
        technical_spec_links=formatted_hyperlinks,
        status=status,
        eligibility_score=score_info["eligibility_score"],
        ai_confidence=ai_confidence,
        human_review_required=needs_review,
        rich=rich,
    )


# --------------------------------------------------------------------------
# Pipeline Driver & Google Sheet Updating
# --------------------------------------------------------------------------

def load_procmart_config(config_path: str) -> dict:
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[WARN] Could not load config file {config_path}: {e}. Proceeding with defaults.", file=sys.stderr)
        return {}


def fetch_sheet_rows_http(sheet_id: str, tab_name: str) -> List[Dict[str, str]]:
    """Fallback HTTP CSV fetch if gspread is not authenticated."""
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={tab_name}"
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(csv_url, headers=headers)
    response.raise_for_status()

    csv_data = io.StringIO(response.text)
    reader = csv.DictReader(csv_data)

    target_items = []
    for row in reader:
        bid_no = row.get("Bid No", "").strip()
        bid_url = row.get("Bid Doc URL", "").strip()
        if bid_no and bid_url:
            target_items.append({"bid_no": bid_no, "bid_url": bid_url, **row})

    return target_items


def update_google_sheet_direct(
    sheet_id: str,
    tab_name: str,
    records: List[BidRecord],
    service_account_file: str = "service_account.json"
):
    """
    Quota-efficient Google Sheet write-back.

    Uses worksheet.batch_update() instead of update_cell() so the
    complete write-back is sent as a single Google Sheets API request.

    IMPORTANT:
    Because batch_update() is called on the selected worksheet object,
    ranges must be plain cell/range addresses such as Q2:Y2.
    Do NOT prepend the worksheet/tab name.
    """

    if not gspread:
        print(
            "[WARN] gspread library not installed. "
            "Install via `pip install gspread` to write back to Google Sheets.",
            file=sys.stderr
        )
        return

    import time

    try:
        # ---------------------------------------------------------
        # 1. Authenticate and select the exact worksheet/tab
        # ---------------------------------------------------------
        gc = gspread.service_account(
            filename=service_account_file
        )

        sh = gc.open_by_key(sheet_id)
        worksheet = sh.worksheet(tab_name)

        print(f"[INFO] Connected to Google Sheet tab: '{tab_name}'")

        # ---------------------------------------------------------
        # 2. Read existing headers once
        # ---------------------------------------------------------
        headers = worksheet.row_values(1)

        target_cols = [
            "BidDoc_eligibility",
            "ChildDoc_eligibility",
            "Procmart_eligibility",
            "technical_spec_links",
            "Eligibility_Status",
            "Eligibility_Score",
            "AI_Confidence",
            "Human_Review_Required",
            "Eligibility_JSON",
        ]

        # ---------------------------------------------------------
        # 3. Add missing headers in ONE request
        # ---------------------------------------------------------
        missing_cols = [
            col for col in target_cols
            if col not in headers
        ]

        if missing_cols:
            start_col = len(headers) + 1
            end_col = start_col + len(missing_cols) - 1

            start_cell = gspread.utils.rowcol_to_a1(1, start_col)
            end_cell = gspread.utils.rowcol_to_a1(1, end_col)

            worksheet.update(
                range_name=f"{start_cell}:{end_cell}",
                values=[missing_cols],
                value_input_option="USER_ENTERED"
            )

            headers.extend(missing_cols)

            print(
                f"[INFO] Added missing headers: "
                f"{', '.join(missing_cols)}"
            )

        # ---------------------------------------------------------
        # 4. Build column-index map
        # ---------------------------------------------------------
        col_indices = {
            col: headers.index(col) + 1
            for col in target_cols
        }

        # ---------------------------------------------------------
        # 5. Locate Bid No column
        # ---------------------------------------------------------
        bid_no_col_idx = (
            headers.index("Bid No") + 1
            if "Bid No" in headers
            else 3
        )

        # ---------------------------------------------------------
        # 6. Read Bid No column once
        # ---------------------------------------------------------
        bid_no_values = worksheet.col_values(bid_no_col_idx)

        bid_row_map = {}

        for row_idx, bid_no in enumerate(bid_no_values, start=1):
            normalized_bid = str(bid_no).strip()
            if normalized_bid:
                bid_row_map.setdefault(normalized_bid, row_idx)

        # ---------------------------------------------------------
        # 7. Prepare batch updates
        # ---------------------------------------------------------
        batch_data = []
        updated_count = 0
        skipped_count = 0

        for rec in records:

            bid_number = str(rec.bid_number).strip()
            row_idx = bid_row_map.get(bid_number)

            if row_idx is None:
                print(
                    f"[WARN] Bid No '{bid_number}' "
                    f"not found in Google Sheet. Skipping."
                )
                skipped_count += 1
                continue

            compact_json = json.dumps(
                rec.rich,
                ensure_ascii=False,
                separators=(",", ":")
            )

            # Google Sheets cell limit is approximately 50,000 chars.
            if len(compact_json) > 49000:
                compact_json = (
                    compact_json[:48980]
                    + '..."(truncated)"}'
                )

            values = {
                "BidDoc_eligibility": rec.bid_doc_eligibility,
                "ChildDoc_eligibility": rec.child_doc_eligibility,
                "Procmart_eligibility": rec.procmart_eligibility,
                "technical_spec_links": rec.technical_spec_links,
                "Eligibility_Status": rec.status,
                "Eligibility_Score": rec.eligibility_score,
                "AI_Confidence": rec.ai_confidence,
                "Human_Review_Required": (
                    "YES" if rec.human_review_required else "NO"
                ),
                "Eligibility_JSON": compact_json,
            }

            for col_name, value in values.items():

                col_idx = col_indices[col_name]

                cell = gspread.utils.rowcol_to_a1(
                    row_idx,
                    col_idx
                )

                # IMPORTANT:
                # Only "Q2", "R2", etc. here.
                # Do NOT use "'2026-08-23'!Q2".
                batch_data.append({
                    "range": cell,
                    "values": [[value]]
                })

            updated_count += 1

        # ---------------------------------------------------------
        # 8. Perform ONE batch write
        # ---------------------------------------------------------
        if batch_data:

            max_retries = 5

            for attempt in range(max_retries):

                try:
                    worksheet.batch_update(
                        batch_data,
                        value_input_option="USER_ENTERED"
                    )
                    break

                except Exception as api_error:

                    error_text = str(api_error)

                    is_quota_error = (
                        "429" in error_text
                        or "Quota exceeded" in error_text
                        or "quota" in error_text.lower()
                    )

                    if (
                        not is_quota_error
                        or attempt == max_retries - 1
                    ):
                        raise

                    wait_seconds = 2 ** attempt

                    print(
                        f"[WARN] Google Sheets quota limit hit. "
                        f"Retrying in {wait_seconds}s "
                        f"(attempt {attempt + 1}/{max_retries})..."
                    )

                    time.sleep(wait_seconds)

        print(
            f"Successfully updated Google Sheet tab '{tab_name}'."
        )
        print(f"[INFO] Rows updated: {updated_count}")
        print(f"[INFO] Rows skipped: {skipped_count}")
        print(f"[INFO] Batch entries: {len(batch_data)}")

    except Exception as e:
        print(
            f"[ERROR] Failed to update Google Sheet directly: {e}",
            file=sys.stderr
        )

def write_local_outputs(records: List[BidRecord], out_csv: str, out_json: str, tab_label: str):
    fieldnames = ["bid_number", "status", "eligibility_score", "ai_confidence", "human_review_required",
                  "BidDoc_eligibility", "ChildDoc_eligibility", "Procmart_eligibility", "technical_spec_links"]
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in records:
            writer.writerow({
                "bid_number": r.bid_number,
                "status": r.status,
                "eligibility_score": r.eligibility_score,
                "ai_confidence": r.ai_confidence,
                "human_review_required": r.human_review_required,
                "BidDoc_eligibility": r.bid_doc_eligibility,
                "ChildDoc_eligibility": r.child_doc_eligibility,
                "Procmart_eligibility": r.procmart_eligibility,
                "technical_spec_links": r.technical_spec_links,
            })
    print(f"Local CSV written to: {out_csv}")

    payload = {
        "generated_at": datetime.now().isoformat(),
        "sheet_tab": tab_label,
        "bid_count": len(records),
        "bids": [r.rich for r in records],
    }
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"Rich JSON (for the dashboard) written to: {out_json}")


def run_pipeline(sheet_id: str, out_csv: str, out_json: str, config_path: str, tab_override: Optional[str],
                  fetch_children: bool, service_account_file: str, write_back: bool):
    config = load_procmart_config(config_path)

    tab_name = tab_override or (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    print(f"Targeting Google Sheet Tab: '{tab_name}'")

    print(f"Fetching Google Sheet records...")
    target_items = fetch_sheet_rows_http(sheet_id, tab_name)
    print(f"Retrieved {len(target_items)} records from Google Sheet.")

    records: List[BidRecord] = []
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})

    for index, item in enumerate(target_items, start=1):
        bid_no = item["bid_no"]
        doc_url = item["bid_url"]
        print(f"[{index}/{len(target_items)}] Processing Bid: {bid_no}...")

        try:
            resp = session.get(doc_url, timeout=30)
            resp.raise_for_status()
            rec = parse_bid_pdf_stream(resp.content, bid_no, doc_url, session, config, item, fetch_children)
            records.append(rec)
        except Exception as e:
            print(f"[WARN] Failed to download or parse {bid_no} ({doc_url}): {e}", file=sys.stderr)
            records.append(BidRecord(
                bid_number=bid_no,
                bid_doc_url=doc_url,
                bid_doc_eligibility=f"*Error: Could not parse document - {e}",
                child_doc_eligibility="*N/A",
                procmart_eligibility="*Status: UNVERIFIED",
                technical_spec_links="",
                status="Disqualified",
                eligibility_score=0,
                ai_confidence=0,
                human_review_required=True,
                rich={
                    "bid_number": bid_no, "bid_doc_url": doc_url,
                    "status": "Disqualified", "status_reason": f"Could not download or parse bid document: {e}",
                    "eligibility_score": 0, "ai_confidence": 0, "human_review_required": True,
                    "human_review_reasons": ["Document fetch/parse failure"],
                    "requirements": [], "documents": {"linked_documents_found": 0, "linked_documents_processed": 0,
                                                        "linked_documents_failed": 0, "children": []},
                    "scoring": {}, "technical_spec_links": "",
                    "human_review": {"resolved": False, "final_status": None, "reviewer_comment": "", "override_reason": ""},
                },
            ))

    write_local_outputs(records, out_csv, out_json, tab_name)

    if write_back:
        update_google_sheet_direct(sheet_id, tab_name, records, service_account_file)
    else:
        print("[INFO] Skipping Google Sheet write-back (--no-writeback set). Local CSV/JSON are up to date.")


# --------------------------------------------------------------------------
# Local/offline test mode (evaluate a folder of already-downloaded PDFs)
# --------------------------------------------------------------------------

def run_local_folder(folder: str, out_csv: str, out_json: str, config_path: str, fetch_children: bool):
    config = load_procmart_config(config_path)
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})

    records: List[BidRecord] = []
    files = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
    for index, fname in enumerate(files, start=1):
        path = os.path.join(folder, fname)
        print(f"[{index}/{len(files)}] Evaluating local file: {fname}")
        with open(path, "rb") as fh:
            pdf_bytes = fh.read()
        bid_no_guess = os.path.splitext(fname)[0]
        try:
            rec = parse_bid_pdf_stream(pdf_bytes, bid_no_guess, path, session, config, {}, fetch_children)
        except Exception as e:
            print(f"[WARN] Failed to evaluate {fname}: {e}", file=sys.stderr)
            rec = BidRecord(bid_number=bid_no_guess, bid_doc_url=path, status="Disqualified",
                             bid_doc_eligibility=f"*Error: {e}", child_doc_eligibility="*N/A",
                             procmart_eligibility="*Status: UNVERIFIED", technical_spec_links="",
                             eligibility_score=0, ai_confidence=0, human_review_required=True,
                             rich={"bid_number": bid_no_guess, "bid_doc_url": path, "status": "Disqualified",
                                   "status_reason": f"Parse error: {e}", "eligibility_score": 0, "ai_confidence": 0,
                                   "human_review_required": True, "human_review_reasons": ["Parse error"],
                                   "requirements": [], "documents": {"linked_documents_found": 0,
                                                                      "linked_documents_processed": 0,
                                                                      "linked_documents_failed": 0, "children": []},
                                   "scoring": {}, "technical_spec_links": "",
                                   "human_review": {"resolved": False, "final_status": None,
                                                     "reviewer_comment": "", "override_reason": ""}})
        records.append(rec)

    write_local_outputs(records, out_csv, out_json, "local-test")


def main():
    default_sheet_id = "1BiZSN_TvhyNfX2aa4G9qucCs1GLxpbsPCWoKnsFuF4s"

    parser = argparse.ArgumentParser(description="GeM Bid Eligibility Extraction Pipeline.")
    parser.add_argument("--sheet-id", default=default_sheet_id, help="Google Sheet ID.")
    parser.add_argument("--date", default=None, help="Sheet tab / date (YYYY-MM-DD) to process. Default: yesterday.")
    parser.add_argument("--out-csv", default="eligibility_output.csv", help="Output CSV file.")
    parser.add_argument(
        "--out-json",
        default=None,
        help="Output rich JSON file. If omitted, automatically uses bids_<date>.json."
    )
    parser.add_argument("--config", default="procmart_config.json", help="Path to procmart_config.json file.")
    parser.add_argument("--service-account", default="service_account.json",
                         help="Path to the Google service-account credentials JSON for Sheet write-back.")
    parser.add_argument("--no-writeback", action="store_true",
                         help="Skip writing results back into the Google Sheet (local CSV/JSON are still produced).")
    parser.add_argument("--no-child-docs", action="store_true", help="Skip downloading/parsing linked child documents.")
    parser.add_argument("--local-folder", default=None,
                         help="Offline test mode: evaluate every PDF in this folder instead of hitting the Google Sheet.")
    args = parser.parse_args()

    # Automatically create a date-specific JSON filename for the dashboard.
    # Example: --date 2026-08-23 -> bids_2026-08-23.json
    if args.out_json is None:
        json_date = (
            args.date
            or (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        )
        args.out_json = f"bids_{json_date}.json"

    if args.local_folder:
        run_local_folder(args.local_folder, args.out_csv, args.out_json, args.config, not args.no_child_docs)
    else:
        run_pipeline(args.sheet_id, args.out_csv, args.out_json, args.config, args.date,
                     not args.no_child_docs, args.service_account, not args.no_writeback)


if __name__ == "__main__":
    main()
