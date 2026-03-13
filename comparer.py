"""
==============================================================================
PDF Permission Comparator: T3 vs PROD
==============================================================================
Description:
    This script extracts text from two PDF files representing system permission
    configurations (T3 / pre-production and PROD / production instances),
    parses the permission entries, and identifies any differences between them.

Requirements:
    pip install pymupdf

Usage:
    1. Place both PDF files in the same directory as this script (or update paths).
    2. Run:  python compare_permissions.py
    3. Review the console output and the generated report file.

Author:  AI Assistant
Date:    2025
==============================================================================
"""

import fitz  # PyMuPDF
import re
import os
import sys
from datetime import datetime
from collections import OrderedDict


# ============================================================================
# CONFIGURATION
# ============================================================================
import glob as _glob

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def find_pdf_pairs():
    """
    Find all T3/PROD PDF pairs in the script's directory or samples/ subdirectory.

    Matches files by their shared prefix before _T3.pdf / _PROD.pdf.
    Returns a list of (t3_path, prod_path, base_name) tuples.
    """
    search_dirs = [SCRIPT_DIR, os.path.join(SCRIPT_DIR, "samples")]
    t3_files = []
    for d in search_dirs:
        t3_files.extend(sorted(_glob.glob(os.path.join(d, "*_T3.pdf"))))

    pairs = []
    for t3_path in t3_files:
        base = t3_path.rsplit("_T3.pdf", 1)[0]
        prod_path = base + "_PROD.pdf"
        base_name = os.path.basename(base)
        if os.path.exists(prod_path):
            pairs.append((t3_path, prod_path, base_name))
        else:
            print(f"[WARNING] No PROD match for: {os.path.basename(t3_path)}")
    if not pairs:
        print("[ERROR] No T3/PROD PDF pairs found in", SCRIPT_DIR)
        sys.exit(1)
    return pairs


# ============================================================================
# PDF TEXT EXTRACTION
# ============================================================================
def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Extract all text content from a PDF file using PyMuPDF.
    Filters out page-break noise (URLs, page headers/footers).

    Args:
        pdf_path: Path to the PDF file.

    Returns:
        Full text content of the PDF as a single string.
    """
    if not os.path.exists(pdf_path):
        print(f"[ERROR] File not found: {pdf_path}")
        sys.exit(1)

    full_text = []
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = page.get_text("blocks")
        page_lines = []
        for b in blocks:
            # b = (x0, y0, x1, y1, text, block_no, block_type)
            if b[-1] != 0:  # skip image blocks
                continue
            # Filter out page header area (y < 30) — contains repeated timestamps and titles
            if b[1] < 30:
                continue
            text = b[4]
            if text.strip():
                page_lines.append(text)
        if page_lines:
            full_text.append("".join(page_lines))
        else:
            print(f"  [WARNING] No text extracted from page {page_num + 1} of {pdf_path}")
    doc.close()

    # Filter out page-break noise lines
    cleaned = []
    for line in "\n".join(full_text).split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        # Skip URL lines (page footers)
        if stripped.startswith("http://") or stripped.startswith("https://"):
            continue
        # Skip page header lines like "3/12/26, 2:19 PM View Role for ..."
        if re.match(r'^\d+/\d+/\d+,\s+\d+:\d+\s+(AM|PM)', stripped):
            continue
        # Skip "View Role for ..." title lines repeated on each page
        if stripped.startswith("View Role for ") or stripped.startswith("View ALL_"):
            continue
        # Skip page number lines like "1/39", "18/26"
        if re.match(r'^\d+/\d+$', stripped):
            continue
        cleaned.append(stripped)

    return "\n".join(cleaned)


# ============================================================================
# SECTION PARSING
# ============================================================================

# Known top-level section headers that appear in the permission documents
SECTION_HEADERS = [
    "Role Name",
    "Role Description",
    "Permission settings",
    "User Type",
    "RBP-Only",
    "Last Modified By",
    "Last Modified Date",
    "Employee Data",
    "ESS",
    "Miscellaneous Permissions",
    "Manage Position",
    "Reports Permission",
    "Manage User",
    "Administrator Permissions",
    "Workflow",
    "MDF Foundation Objects",
    "Metadata Framework",
    "Employee Central Effective Dated Entities",
    "Employee Central Object-Level Permissions",
    "Employee Central Import and Export Data Permissions",
]


def _is_pure_perm_line(text: str) -> bool:
    """Return True if the line contains ONLY permission keywords and | separators.

    Examples that return True:
        "Correct | Delete"
        "View | Edit"
        "Import/Export"
    Examples that return False:
        "First Name View | Edit"
        "Compensation Information > Correct | Delete"
    """
    compound_perms = {"View Current", "View History", "Edit/Insert", "Import/Export"}
    single_perms = {"View", "Edit", "Insert", "Delete", "Import", "Export",
                    "Approve", "Create", "Correct", "None", "Admin", "Yes", "No"}

    clean = text.rstrip(" ★")
    # Strip out compound perms first
    for cp in compound_perms:
        clean = clean.replace(cp, " ")
    # Split by | and check every part
    parts = [p.strip() for p in clean.split("|")]
    for part in parts:
        words = part.split()
        for w in words:
            if w not in single_perms:
                return False
    # Must have at least one token
    return bool(text.strip())


def parse_sections(text: str, element_headers=None) -> OrderedDict:
    """
    Parse the extracted PDF text into logical sections and subsections.

    Builds hierarchical keys like "Employee Data > Employee Profile > Country"
    so that subsections (marked with †) are preserved in the comparison.

    Lines ending with † or ⁜ (with no permission keywords) are treated as
    subsection headers. Lines containing permission keywords (View, Edit, etc.)
    are treated as permission entries.

    Args:
        text: Full extracted text from a PDF.

    Returns:
        OrderedDict mapping section names to their content lines.
        Content lines use the format: "Subsection > Field  permissions"
    """
    raw_lines = text.split("\n")
    sections = OrderedDict()
    current_section = "__HEADER__"
    current_subsection = ""
    sections[current_section] = []

    lines = [l.strip() for l in raw_lines if l.strip()]

    for idx, stripped in enumerate(lines):

        # Stop parsing once we hit the "Assignments" section (after Metadata Framework)
        # Also stop at "ID:" lines that precede the Assignments heading
        if stripped.lower().startswith("assignments"):
            break
        if current_section == "Metadata Framework" and stripped.startswith("ID:"):
            break

        # Check if this line is a known top-level section header
        matched_section = None
        for header in SECTION_HEADERS:
            if stripped.lower().startswith(header.lower()):
                matched_section = header
                break

        if matched_section:
            current_section = matched_section
            current_subsection = ""
            if current_section not in sections:
                sections[current_section] = []
            remainder = stripped[len(matched_section):].strip(": ").strip()
            if remainder:
                sections[current_section].append(remainder)
            continue

        # Check if this is a subsection header (ends with † or ⁜).
        # The † / ⁜ marker is the definitive signal for a subsection header.
        # A † line is NOT a subsection header only if the NEXT line is also
        # a † line (indicating a chain of standalone toggles, e.g.,
        # "Delete Assigned Team Goals †" followed by "Assign Team Goals †").
        clean = stripped.rstrip(" †⁜★")
        if stripped.endswith("†") or stripped.endswith("⁜"):
            # A † line is a subsection header if it has children.
            # Look ahead: if the next non-empty line is also a † / ⁜ line
            # or is a known section header, then this † line is a
            # standalone toggle (e.g., "Delete Assigned Team Goals †"),
            # not a subsection header.
            next_idx = idx + 1
            while next_idx < len(lines) and not lines[next_idx].strip():
                next_idx += 1
            next_line = lines[next_idx].strip() if next_idx < len(lines) else ""
            next_is_dagger = "†" in next_line or next_line.endswith("⁜")
            next_is_section = any(
                next_line.lower().startswith(h.lower()) for h in SECTION_HEADERS
            )
            next_is_element = (
                element_headers and next_line in element_headers
            )
            if next_is_dagger or next_is_section or next_is_element:
                # This † line is a standalone toggle — treat as a regular
                # line, not a subsection header. Fall through to add it.
                pass
            else:
                current_subsection = clean
                continue

        # Build the line to add (with or without subsection prefix)
        if _is_pure_perm_line(stripped):
            # Pure permission continuation (e.g., "Correct | Delete") —
            # don't prefix with subsection so it merges in parse_permission_lines
            add_line = stripped
        elif current_subsection:
            add_line = f"{current_subsection} > {stripped}"
        else:
            add_line = stripped

        # If the previous line in this section ended with "|", this line
        # is a continuation (permissions and/or name wrapped across a
        # page break).  Merge it into the previous line.
        sec_lines = sections.setdefault(current_section, [])
        if sec_lines and sec_lines[-1].rstrip().endswith("|"):
            # Strip duplicate subsection prefix if present on continuation
            if current_subsection and stripped.startswith(current_subsection + " > "):
                stripped = stripped[len(current_subsection) + 3:]
            sec_lines[-1] = sec_lines[-1].rstrip() + " " + stripped
        else:
            sec_lines.append(add_line)

    return sections


# ============================================================================
# PERMISSION LINE PARSING
# ============================================================================

def parse_permission_lines(lines: list) -> OrderedDict:
    """
    Parse a list of content lines into individual permission entries.

    Handles compound permissions like "View Current | View History | Edit/Insert"
    and subsection-prefixed lines like "Employee Profile > Country View | Edit".

    Permission keywords are only extracted when they appear in a clear
    permission context: after a | separator, or as trailing words at the
    end of a line. Words like "Import" or "View" embedded in the middle
    of a name (e.g., "Employee Central Import Entities") are kept as
    part of the name.

    Lines with no permission syntax are treated as standalone enabled
    items (e.g., admin capabilities like "Rehire Inactive Employee").

    Multi-line field names (continuation lines with no permissions) are
    joined to the previous entry's name.

    Args:
        lines: List of text lines from a section.

    Returns:
        OrderedDict mapping permission/item name -> set of granted rights.
    """
    permissions = OrderedDict()

    # Compound permission phrases (order matters — check longer phrases first)
    compound_perms = [
        "View Current", "View History", "Edit/Insert", "Import/Export",
    ]
    # Single permission keywords
    single_perms = {
        "View", "Edit", "Insert", "Delete", "Import", "Export",
        "Approve", "Create", "Correct", "None", "Admin", "Yes", "No",
    }

    def _has_perm_syntax(text):
        """Check if a line has explicit permission syntax (| separator or
        trailing permission keywords)."""
        if "|" in text:
            return True
        clean = text.rstrip(" ★")
        words = clean.split()
        if not words:
            return False
        for cp in compound_perms:
            if clean.endswith(cp):
                return True
        return words[-1] in single_perms

    def extract_name_and_perms(text):
        """Split a line into (field_name, set_of_permissions)."""
        # Remove trailing ★ markers
        text = text.rstrip(" ★")

        # If no explicit permission syntax, treat as a standalone item
        if not _has_perm_syntax(text):
            return text.strip(), {"Enabled"}

        # First extract compound permissions from the text
        granted = set()
        remaining = text
        for cp in compound_perms:
            while cp in remaining:
                granted.add(cp)
                remaining = remaining.replace(cp, " ", 1)

        # Split what's left by | to find single-word permissions
        parts = [p.strip() for p in remaining.split("|")]
        name_parts = []
        for part in parts:
            words = part.split()
            part_name_words = []
            for w in words:
                if w in single_perms:
                    granted.add(w)
                else:
                    part_name_words.append(w)
            if part_name_words:
                name_parts.append(" ".join(part_name_words))

        name = " ".join(name_parts).strip()
        return name, granted

    last_name = None
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Lines ending with a parenthetical identifier like "(TER_FRA_MA_SD)"
        # belong to the previous entry's name.  This covers:
        #   - Bare IDs:        "(TER_ENDTC)"
        #   - Prefixed IDs:    "Subsection > (TER_FRA_MA_SD)"
        #   - IDs with text:   "art.30§1pkt.5 (TER_POL_68)"
        # Strip the subsection prefix first, then grab everything including
        # the text before the ID as the suffix to append.
        line_no_prefix = stripped
        if ">" in stripped:
            line_no_prefix = stripped.split(">", 1)[1].strip()
        # Must contain at least one underscore to distinguish real IDs
        # like (TER_ENDTC) from normal parentheticals like (All) or (Employee)
        id_match = re.search(r'\(([A-Za-z0-9\-]*_[A-Za-z0-9_\-]*)\)\s*$', line_no_prefix)
        if (id_match and not _has_perm_syntax(stripped) and last_name is not None):
            # Only merge as continuation if the previous entry doesn't already
            # have real permissions (View, Edit, etc.).  If it does, this line
            # is a new entry that happens to contain an ID — not a continuation.
            prev_perms = permissions.get(last_name, set())
            real_perms = prev_perms - {"Enabled"}
            if not real_perms:
                suffix = line_no_prefix.strip()
                prev_perms = permissions.pop(last_name)
                new_name = f"{last_name} {suffix}"
                permissions[new_name] = prev_perms
                last_name = new_name
                continue

        name, granted = extract_name_and_perms(stripped)

        if not name and not granted:
            continue

        # If the line has only permissions and no name, it's a continuation
        # of the previous entry's permissions (e.g., "Correct | Delete" on
        # a new line after "First Name View Current | Edit/Insert |")
        if not name and granted and last_name is not None:
            permissions[last_name] |= granted
            continue

        # If the line has no permissions (just "Enabled"), it might be a
        # continuation of the previous field name (multi-line wrapping).
        # BUT only merge if the current line looks like a name fragment
        # (no ">" subsection separator, starts lowercase or is very short).
        # Otherwise it's a standalone enabled item.
        if granted == {"Enabled"} and last_name is not None:
            prev_perms = permissions[last_name]
            is_name_fragment = (
                ">" not in name
                and not name[0].isupper() if name else False
            )
            if prev_perms == {"Enabled"} and is_name_fragment:
                permissions.pop(last_name)
                new_name = f"{last_name} {name}"
                permissions[new_name] = prev_perms
                last_name = new_name
                continue

        if not name:
            name = stripped

        # Handle duplicate names
        base_name = name
        counter = 1
        while name in permissions:
            counter += 1
            name = f"{base_name} (#{counter})"
        permissions[name] = granted
        last_name = name

    return permissions


# ============================================================================
# PYMUPDF STRUCTURED EXTRACTION
# ============================================================================

# Layout thresholds (x-positions) based on the PDF visual structure.
# These classify each line by its indentation level.
_X_SECTION = 82      # x < 82 → top-level section (bold), e.g., "User Permissions"
_X_ELEMENT = 100     # 82 ≤ x < 100 → element header (bold), e.g., "Calibration"
_X_SUBSECTION = 135  # 100 ≤ x < 135 → subsection / object name, e.g., "Job Information †"
# Fields are at x ≈ 149, permission values at x ≈ 326+
_X_PERM = 300        # x ≥ 300 → permission values, e.g., "View Current | View History"


def extract_pdf_structured(pdf_path):
    """
    Extract permissions from a PDF using PyMuPDF's layout information.

    Uses x-position and bold detection to determine hierarchy:
        x ≈ 77  bold  → top-level section ("User Permissions", "Administrator Permissions")
        x ≈ 89  bold  → element header ("Calibration", "Goals", "Admin Center Permissions")
        x ≈ 119       → subsection or standalone item ("Job Information †", "Goal Management Access")
        x ≈ 149       → field name ("First Name", "Object-Level Permissions")
        x ≈ 326+      → permission value ("View Current | View History")

    Returns:
        List of dicts: [{section, element, subsection, field, permissions_str}, ...]
    """
    doc = fitz.open(pdf_path)
    entries = []

    # First pass: collect all visual lines across pages
    visual_lines = []  # [(x, y, is_bold, text, page_num), ...]

    for pg_num in range(len(doc)):
        page = doc[pg_num]
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if "lines" not in b:
                continue
            for line in b["lines"]:
                x_pos = None
                y_pos = None
                is_bold = False
                texts = []
                for span in line["spans"]:
                    t = span["text"]
                    if x_pos is None and t.strip():
                        x_pos = span["origin"][0]
                        y_pos = span["origin"][1]
                    if "Bold" in span.get("font", ""):
                        is_bold = True
                    texts.append(t)
                full = "".join(texts).strip()
                if not full or x_pos is None:
                    continue
                # Skip page headers/footers
                if full.startswith("http://") or full.startswith("https://"):
                    continue
                if re.match(r"^\d+/\d+/\d+,\s+\d+:\d+\s+(AM|PM)", full):
                    continue
                if re.match(r"^\d+/\d+$", full):  # "18/26" page numbers
                    continue
                # Skip legend lines
                if full.startswith("★=") or full.startswith("⁜=") or full.startswith("†="):
                    continue
                # Skip page title repeated on each page
                if "View Role for" in full or full.startswith("View ALL_"):
                    continue
                # Skip the role metadata header area on the first page
                if pg_num == 0 and y_pos < 220:
                    continue
                visual_lines.append((x_pos, y_pos, is_bold, full, pg_num))

    doc.close()

    # Sort visual lines by (page, y-position, x-position) so that field names
    # and their permission values on the same y-line appear together, and
    # wrapped continuations at higher y-positions come after.
    visual_lines.sort(key=lambda v: (v[4], v[1], v[0]))

    # Second pass: pair field names (x < _X_PERM) with their permission values
    # (x ≥ _X_PERM) that appear on the same y-position or immediately after.
    #
    # Build "logical rows": each is a (level, is_bold, name_text, perm_text) tuple.
    logical_rows = []
    i = 0
    while i < len(visual_lines):
        x, y, bold, text, pg = visual_lines[i]

        if x >= _X_PERM:
            # Standalone permission text (continuation of previous row's perms)
            if logical_rows:
                prev = logical_rows[-1]
                logical_rows[-1] = (prev[0], prev[1], prev[2],
                                    (prev[3] + " " + text).strip() if prev[3] else text)
            i += 1
            continue

        # Determine level from x position
        if x < _X_SECTION:
            level = "section"
        elif x < _X_ELEMENT:
            level = "element"
        elif x < _X_SUBSECTION:
            level = "subsection"
        else:
            level = "field"

        # Collect the permission text from the same line or next lines at x ≥ _X_PERM
        perm_text = ""
        j = i + 1

        # Check if the next visual line(s) at x ≥ _X_PERM are on roughly the
        # same y-position (same line) or are continuation lines (perms that
        # wrapped to the next line).
        while j < len(visual_lines):
            nx, ny, nb, nt, npg = visual_lines[j]
            if nx >= _X_PERM:
                # Permission value (same line or continuation)
                perm_text = (perm_text + " " + nt).strip() if perm_text else nt
                j += 1
            else:
                break

        # Check if the next non-perm lines at the same level are continuations
        # of the current name (wrapped text). Works for both field-level (x≈149)
        # and subsection-level (x≈119) wrapping.
        # E.g., "PRT - Continued Sickness Pay" (x=149, no perms)
        #        "Period" (x=149, no perms)
        #        "View Current | ..." (x=326, perms)
        # Or:   "Goal Plan Permissions †(Learning Activity, ..." (x=119)
        #        "2025, Business Goals ...)" (x=119)
        name_text = text
        if not perm_text:
            # Determine the x-range for same-level continuation lines
            if level == "field":
                x_lo, x_hi = _X_SUBSECTION, _X_PERM
            elif level == "subsection":
                x_lo, x_hi = _X_ELEMENT, _X_SUBSECTION
            else:
                x_lo, x_hi = -1, -1  # no wrapping for section/element

            def _is_name_continuation(prev_text, next_text):
                """Check if next_text is a continuation of prev_text."""
                nt_stripped = next_text.strip()
                if not nt_stripped:
                    return False
                # Starts with lowercase, digit, comma, or closing paren
                if nt_stripped[0].islower() or nt_stripped[0].isdigit():
                    return True
                if nt_stripped.startswith((",", ")", "-", "(")):
                    return True
                # Previous text has unbalanced parens
                if prev_text.count("(") > prev_text.count(")"):
                    return True
                # Previous text ends with " -" or " –" (line wrapped mid-name)
                pt_stripped = prev_text.rstrip()
                if pt_stripped.endswith("-") or pt_stripped.endswith("–") or pt_stripped.endswith("of"):
                    return True
                return False

            while j < len(visual_lines):
                nx, _, nb, nt, _ = visual_lines[j]
                if nx >= _X_PERM:
                    # Found perms for this wrapped entry
                    perm_text = nt
                    j += 1
                    while j < len(visual_lines) and visual_lines[j][0] >= _X_PERM:
                        perm_text = perm_text + " " + visual_lines[j][3]
                        j += 1
                    break
                elif x_lo <= nx < x_hi and not nb and _is_name_continuation(name_text, nt):
                    # Same-level continuation fragment
                    name_text = name_text + " " + nt
                    j += 1
                else:
                    break

        logical_rows.append((level, bold, name_text, perm_text))
        i = j if j > i + 1 else i + 1

    # Post-process: merge continuation rows into the previous row.
    # Case 1: Short field names (≤2 words) with no perms after a row WITH perms
    #   e.g., "ARG - Continued Sickness Pay" + "Period"
    # Case 2: Field rows with no perms whose previous row also has no perms
    #   e.g., "INVOL - Antecipation End of" + "Contract - Company (TER_BRA_BH)"
    #   These chain together and eventually the next row with perms gets them.
    # Case 3: Field rows with no perms where the previous row's name ends with
    #   a hyphen/dash — indicates a wrapped name regardless of word count.
    merged_rows = []
    for row in logical_rows:
        level, bold, name, perms = row
        if (merged_rows
                and not perms
                and not bold
                and level == "field"
                and merged_rows[-1][0] == level  # same level
                and not merged_rows[-1][1]):  # previous not bold
            prev = merged_rows[-1]
            # A no-perms, non-bold field row is always a continuation of the
            # previous field row (wrapped long name in PDF)
            merged_rows[-1] = (prev[0], prev[1], prev[2] + " " + name, prev[3] or perms)
            continue
        merged_rows.append(row)
    logical_rows = merged_rows

    # Third pass: build hierarchical entries from logical rows
    current_section = ""
    current_element = ""
    current_subsection = ""

    # Stop keywords — stop parsing when we hit these sections
    stop_sections = {"Assignments"}

    for idx, (level, bold, name, perms) in enumerate(logical_rows):
        clean_name = name.rstrip(" †⁜★").strip()

        if level == "section":
            if clean_name in stop_sections:
                break
            current_section = clean_name
            current_element = ""
            current_subsection = ""
            continue

        if level == "element" and bold:
            current_element = clean_name
            current_subsection = ""
            continue

        if level == "subsection":
            # Subsection / object name at x≈119
            if perms:
                # Has perms on the same line — it's a standalone permission entry.
                entries.append({
                    "section": current_section,
                    "element": current_element,
                    "subsection": current_subsection,
                    "field": name,
                    "permissions_str": perms,
                })
            else:
                # No perms — check if the next row is a child (field level).
                # If so, this is a subsection header that groups those fields.
                # If not, it's a standalone enabled item.
                next_is_child = (
                    idx + 1 < len(logical_rows)
                    and logical_rows[idx + 1][0] == "field"
                )
                if next_is_child:
                    # This is a subsection header that groups field-level children
                    current_subsection = clean_name
                else:
                    # Standalone enabled item at subsection level
                    entries.append({
                        "section": current_section,
                        "element": current_element,
                        "subsection": current_subsection,
                        "field": name,
                        "permissions_str": "Enabled",
                    })
            continue

        if level == "field":
            entries.append({
                "section": current_section,
                "element": current_element,
                "subsection": current_subsection,
                "field": name,
                "permissions_str": perms if perms else "Enabled",
            })

    return entries


# ============================================================================
# COMPARISON ENGINE
# ============================================================================

def compare_permissions(t3_sections: OrderedDict, 
                        prod_sections: OrderedDict) -> list:
    """
    Compare permission sections between T3 and PROD.
    
    Args:
        t3_sections:   Parsed sections from the T3 PDF.
        prod_sections: Parsed sections from the PROD PDF.
    
    Returns:
        List of difference dictionaries with keys:
            - section: Section name
            - item: Permission item name
            - type: 'MISSING_IN_PROD' | 'MISSING_IN_T3' | 'PERMISSION_MISMATCH' | 'SECTION_MISSING_PROD' | 'SECTION_MISSING_T3'
            - t3_value: Value/permissions in T3
            - prod_value: Value/permissions in PROD
    """
    differences = []

    # Skip non-permission metadata sections
    skip_sections = {"__HEADER__", "User Type", "Role Name", "Role Description",
                     "RBP-Only", "Last Modified By", "Last Modified Date"}

    all_sections = list(OrderedDict.fromkeys(
        list(t3_sections.keys()) + list(prod_sections.keys())
    ))

    for section in all_sections:
        if section in skip_sections:
            continue
        t3_lines = t3_sections.get(section, None)
        prod_lines = prod_sections.get(section, None)

        # --- Section-level differences ---
        if t3_lines is None and prod_lines is not None:
            differences.append({
                "section": section,
                "item": "(entire section)",
                "type": "SECTION_MISSING_IN_T3",
                "t3_value": "N/A",
                "prod_value": "; ".join(prod_lines[:5]) + ("..." if len(prod_lines) > 5 else "")
            })
            continue

        if prod_lines is None and t3_lines is not None:
            differences.append({
                "section": section,
                "item": "(entire section)",
                "type": "SECTION_MISSING_IN_PROD",
                "t3_value": "; ".join(t3_lines[:5]) + ("..." if len(t3_lines) > 5 else "")  ,
                "prod_value": "N/A"
            })
            continue

        # --- Line-level / Permission-level comparison ---
        t3_perms = parse_permission_lines(t3_lines)
        prod_perms = parse_permission_lines(prod_lines)

        all_items = list(OrderedDict.fromkeys(
            list(t3_perms.keys()) + list(prod_perms.keys())
        ))

        for item in all_items:
            t3_val = t3_perms.get(item, None)
            prod_val = prod_perms.get(item, None)

            if t3_val is not None and prod_val is None:
                differences.append({
                    "section": section,
                    "item": item,
                    "type": "MISSING_IN_PROD",
                    "t3_value": ", ".join(sorted(t3_val)) if t3_val else "(present, no flags)",
                    "prod_value": "N/A"
                })
            elif prod_val is not None and t3_val is None:
                differences.append({
                    "section": section,
                    "item": item,
                    "type": "MISSING_IN_T3",
                    "t3_value": "N/A",
                    "prod_value": ", ".join(sorted(prod_val)) if prod_val else "(present, no flags)"
                })
            elif t3_val != prod_val:
                differences.append({
                    "section": section,
                    "item": item,
                    "type": "PERMISSION_MISMATCH",
                    "t3_value": ", ".join(sorted(t3_val)) if t3_val else "(none)",
                    "prod_value": ", ".join(sorted(prod_val)) if prod_val else "(none)"
                })

    return differences


# ============================================================================
# RAW TEXT COMPARISON (Fallback / supplementary)
# ============================================================================

def compare_raw_lines(t3_text: str, prod_text: str) -> dict:
    """
    Perform a raw line-by-line comparison as a supplementary check.
    
    Returns:
        Dictionary with 'only_in_t3' and 'only_in_prod' line sets.
    """
    t3_lines = set(line.strip() for line in t3_text.split("\n") if line.strip())
    prod_lines = set(line.strip() for line in prod_text.split("\n") if line.strip())

    only_in_t3 = sorted(t3_lines - prod_lines)
    only_in_prod = sorted(prod_lines - t3_lines)

    return {
        "only_in_t3": only_in_t3,
        "only_in_prod": only_in_prod
    }


# ============================================================================
# REPORT GENERATION
# ============================================================================

def generate_report(differences: list, raw_diff: dict, output_path: str,
                    t3_path: str = "", prod_path: str = "",
                    t3_sections: OrderedDict = None,
                    prod_sections: OrderedDict = None):
    """
    Generate a formatted text report of all identified differences.
    """
    separator = "=" * 90
    sub_separator = "-" * 90
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    report_lines = [
        separator,
        "  PERMISSION COMPARISON REPORT: T3 (Pre-Production) vs PROD (Production)",
        separator,
        f"  Generated: {timestamp}",
        f"  T3 File:   {os.path.basename(t3_path)}",
        f"  PROD File: {os.path.basename(prod_path)}",
        separator,
        "",
    ]

    # ---- Structured Differences ----
    report_lines.append("SECTION 1: STRUCTURED PERMISSION DIFFERENCES")
    report_lines.append(sub_separator)

    if not differences:
        report_lines.append("  *** No structured differences found. Permissions appear identical. ***")
    else:
        report_lines.append(f"  Total differences found: {len(differences)}")
        report_lines.append("")

        # Group by type
        for diff_type in ["PERMISSION_MISMATCH", "MISSING_IN_PROD", "MISSING_IN_T3",
                          "SECTION_MISSING_IN_PROD", "SECTION_MISSING_IN_T3"]:
            type_diffs = [d for d in differences if d["type"] == diff_type]
            if not type_diffs:
                continue

            type_label = {
                "PERMISSION_MISMATCH":    "PERMISSION MISMATCHES (different rights granted)",
                "MISSING_IN_PROD":        "ITEMS PRESENT IN T3 BUT MISSING IN PROD",
                "MISSING_IN_T3":          "ITEMS PRESENT IN PROD BUT MISSING IN T3",
                "SECTION_MISSING_IN_PROD": "ENTIRE SECTIONS MISSING IN PROD",
                "SECTION_MISSING_IN_T3":  "ENTIRE SECTIONS MISSING IN T3",
            }.get(diff_type, diff_type)

            report_lines.append(f"\n  --- {type_label} ({len(type_diffs)} items) ---\n")

            for i, diff in enumerate(type_diffs, 1):
                report_lines.append(f"    {i}. Section : {diff['section']}")
                report_lines.append(f"       Item    : {diff['item']}")
                report_lines.append(f"       T3      : {diff['t3_value']}")
                report_lines.append(f"       PROD    : {diff['prod_value']}")
                report_lines.append("")

    report_lines.append("")
    report_lines.append(separator)

    # ---- Raw Line Differences ----
    report_lines.append("SECTION 2: RAW LINE DIFFERENCES (Supplementary)")
    report_lines.append(sub_separator)

    only_t3 = raw_diff.get("only_in_t3", [])
    only_prod = raw_diff.get("only_in_prod", [])

    report_lines.append(f"  Lines only in T3:   {len(only_t3)}")
    report_lines.append(f"  Lines only in PROD: {len(only_prod)}")
    report_lines.append("")

    if only_t3:
        report_lines.append("  --- Lines ONLY in T3 (not found in PROD) ---")
        for line in only_t3:
            report_lines.append(f"    [T3]  {line}")
        report_lines.append("")

    if only_prod:
        report_lines.append("  --- Lines ONLY in PROD (not found in T3) ---")
        for line in only_prod:
            report_lines.append(f"    [PROD]  {line}")
        report_lines.append("")

    if not only_t3 and not only_prod:
        report_lines.append("  *** No raw line differences found. Files are textually identical. ***")

    # ---- Debug: Parsed Lines per Section ----
    if t3_sections or prod_sections:
        report_lines.append("")
        report_lines.append(separator)
        report_lines.append("SECTION 3: DEBUG — PARSED LINES PER SECTION")
        report_lines.append(sub_separator)

        for label, secs in [("T3", t3_sections), ("PROD", prod_sections)]:
            if not secs:
                continue
            report_lines.append(f"\n  === {label} ===\n")
            for section_name, section_lines in secs.items():
                report_lines.append(f"  [{section_name}]")
                for sl in section_lines:
                    report_lines.append(f"    {sl}")
                report_lines.append("")

    report_lines.append("")
    report_lines.append(separator)
    report_lines.append("  END OF REPORT")
    report_lines.append(separator)

    # Write to file
    report_content = "\n".join(report_lines)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(report_content)

    return report_content


# ============================================================================
# HTML REPORT GENERATION
# ============================================================================

def _html_escape(text):
    """Escape HTML special characters."""
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))


def generate_html_report(differences: list, raw_diff: dict, output_path: str,
                         t3_path: str = "", prod_path: str = ""):
    """Generate a colour-coded HTML report of permission differences."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    t3_name = _html_escape(os.path.basename(t3_path))
    prod_name = _html_escape(os.path.basename(prod_path))

    type_labels = {
        "PERMISSION_MISMATCH":     ("Permission Mismatches", "#e67e22", "Different rights granted between T3 and PROD"),
        "MISSING_IN_PROD":         ("Only in T3", "#e74c3c", "Items present in T3 but missing in PROD"),
        "MISSING_IN_T3":           ("Only in PROD", "#2ecc71", "Items present in PROD but missing in T3"),
        "SECTION_MISSING_IN_PROD": ("Sections only in T3", "#e74c3c", "Entire sections missing in PROD"),
        "SECTION_MISSING_IN_T3":   ("Sections only in PROD", "#2ecc71", "Entire sections missing in T3"),
    }

    # --- Summary counts ---
    counts = {}
    for dt in type_labels:
        counts[dt] = len([d for d in differences if d["type"] == dt])

    only_t3 = raw_diff.get("only_in_t3", [])
    only_prod = raw_diff.get("only_in_prod", [])

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Permission Comparison: T3 vs PROD</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: #f5f6fa; color: #2c3e50; line-height: 1.5; }}
  .container {{ max-width: 1100px; margin: 0 auto; padding: 24px; }}
  header {{ background: linear-gradient(135deg, #2c3e50, #3498db); color: white; padding: 32px; border-radius: 12px; margin-bottom: 24px; }}
  header h1 {{ font-size: 1.6em; margin-bottom: 8px; }}
  header .meta {{ opacity: 0.85; font-size: 0.9em; }}
  header .meta span {{ display: inline-block; margin-right: 24px; }}
  .summary {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 28px; }}
  .card {{ background: white; border-radius: 10px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); text-align: center; border-top: 4px solid #bdc3c7; }}
  .card .num {{ font-size: 2.2em; font-weight: 700; }}
  .card .label {{ font-size: 0.85em; color: #7f8c8d; margin-top: 4px; }}
  .section {{ background: white; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 20px; overflow: hidden; }}
  .section-header {{ padding: 16px 20px; font-weight: 600; font-size: 1.05em; cursor: pointer; display: flex; justify-content: space-between; align-items: center; }}
  .section-header .badge {{ background: rgba(255,255,255,0.3); padding: 2px 10px; border-radius: 12px; font-size: 0.85em; }}
  .section-body {{ padding: 0; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.9em; }}
  th {{ background: #f8f9fa; padding: 10px 16px; text-align: left; font-weight: 600; color: #555; border-bottom: 2px solid #eee; }}
  td {{ padding: 10px 16px; border-bottom: 1px solid #f0f0f0; vertical-align: top; }}
  tr:hover td {{ background: #fafbfc; }}
  .tag {{ display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 0.8em; font-weight: 500; }}
  .tag-t3 {{ background: #fdecea; color: #c0392b; }}
  .tag-prod {{ background: #e8f8f0; color: #27ae60; }}
  .tag-both {{ background: #fef5e7; color: #d35400; }}
  .raw-section {{ padding: 16px 20px; }}
  .raw-section h3 {{ margin-bottom: 10px; font-size: 0.95em; color: #555; }}
  .raw-line {{ font-family: 'SF Mono', 'Consolas', monospace; font-size: 0.82em; padding: 3px 8px; margin: 2px 0; border-radius: 4px; }}
  .raw-t3 {{ background: #fdecea; }}
  .raw-prod {{ background: #e8f8f0; }}
  .identical {{ text-align: center; padding: 40px; color: #27ae60; font-size: 1.2em; }}
  details summary {{ cursor: pointer; padding: 8px 0; }}
  .filter-bar {{ background: white; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); padding: 16px 20px; margin-bottom: 20px; display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }}
  .filter-bar label {{ font-weight: 600; font-size: 0.9em; color: #555; white-space: nowrap; }}
  .filter-bar input {{ flex: 1; min-width: 200px; padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.9em; }}
  .filter-bar input:focus {{ outline: none; border-color: #3498db; box-shadow: 0 0 0 2px rgba(52,152,219,0.2); }}
  .filter-btn {{ padding: 6px 14px; border: 1px solid #ddd; border-radius: 6px; background: #f8f9fa; cursor: pointer; font-size: 0.82em; color: #555; transition: all 0.15s; }}
  .filter-btn:hover {{ background: #e8e8e8; }}
  .filter-btn.active {{ background: #3498db; color: white; border-color: #3498db; }}
  .hidden-row {{ display: none; }}
  .filter-info {{ font-size: 0.82em; color: #95a5a6; margin-left: auto; white-space: nowrap; }}
</style>
</head>
<body>
<div class="container">
  <header>
    <h1>Permission Comparison Report</h1>
    <div class="meta">
      <span>T3: {t3_name}</span>
      <span>PROD: {prod_name}</span>
      <span>Generated: {timestamp}</span>
    </div>
  </header>

  <div class="summary">
    <div class="card" style="border-top-color: #3498db;">
      <div class="num">{len(differences)}</div>
      <div class="label">Total Differences</div>
    </div>
    <div class="card" style="border-top-color: #e67e22;">
      <div class="num">{counts.get("PERMISSION_MISMATCH", 0)}</div>
      <div class="label">Mismatches</div>
    </div>
    <div class="card" style="border-top-color: #e74c3c;">
      <div class="num">{counts.get("MISSING_IN_PROD", 0) + counts.get("SECTION_MISSING_IN_PROD", 0)}</div>
      <div class="label">Only in T3</div>
    </div>
    <div class="card" style="border-top-color: #2ecc71;">
      <div class="num">{counts.get("MISSING_IN_T3", 0) + counts.get("SECTION_MISSING_IN_T3", 0)}</div>
      <div class="label">Only in PROD</div>
    </div>
  </div>
"""

    # --- Filter bar ---
    html += """
  <div class="filter-bar">
    <label>Filter rows:</label>
    <input type="text" id="filterInput" placeholder="Type to exclude rows containing this text...">
    <button class="filter-btn" data-filter="_SD)" onclick="togglePreset(this)">Hide _SD)</button>
    <button class="filter-btn" onclick="clearFilter()">Clear</button>
    <span class="filter-info" id="filterInfo"></span>
  </div>
"""

    if not differences:
        html += '  <div class="identical">No differences found &mdash; permissions are identical.</div>\n'
    else:
        for diff_type in ["PERMISSION_MISMATCH", "MISSING_IN_PROD", "MISSING_IN_T3",
                          "SECTION_MISSING_IN_PROD", "SECTION_MISSING_IN_T3"]:
            type_diffs = [d for d in differences if d["type"] == diff_type]
            if not type_diffs:
                continue

            label, color, desc = type_labels[diff_type]

            html += f"""
  <div class="section">
    <div class="section-header" style="background: {color}; color: white;">
      {_html_escape(label)}: {_html_escape(desc)}
      <span class="badge">{len(type_diffs)}</span>
    </div>
    <div class="section-body">
      <table>
        <tr><th>#</th><th>Section</th><th>Item</th><th>T3</th><th>PROD</th></tr>
"""
            for i, diff in enumerate(type_diffs, 1):
                html += f"""        <tr>
          <td>{i}</td>
          <td>{_html_escape(diff['section'])}</td>
          <td>{_html_escape(diff['item'])}</td>
          <td><span class="tag tag-t3">{_html_escape(diff['t3_value'])}</span></td>
          <td><span class="tag tag-prod">{_html_escape(diff['prod_value'])}</span></td>
        </tr>
"""
            html += """      </table>
    </div>
  </div>
"""

    # --- Raw differences (collapsible) ---
    if only_t3 or only_prod:
        html += """
  <div class="section">
    <div class="section-header" style="background: #95a5a6; color: white;">
      Raw Line Differences (Supplementary)
      <span class="badge">{}</span>
    </div>
    <div class="raw-section">
""".format(len(only_t3) + len(only_prod))

        if only_t3:
            html += f'      <details><summary><strong>Lines only in T3 ({len(only_t3)})</strong></summary>\n'
            for line in only_t3:
                html += f'        <div class="raw-line raw-t3">{_html_escape(line)}</div>\n'
            html += '      </details>\n'

        if only_prod:
            html += f'      <details><summary><strong>Lines only in PROD ({len(only_prod)})</strong></summary>\n'
            for line in only_prod:
                html += f'        <div class="raw-line raw-prod">{_html_escape(line)}</div>\n'
            html += '      </details>\n'

        html += """    </div>
  </div>
"""

    html += """</div>
<script>
const filterInput = document.getElementById('filterInput');
const filterInfo = document.getElementById('filterInfo');

function applyFilter() {
  const term = filterInput.value.trim().toLowerCase();
  const rows = document.querySelectorAll('.section-body table tr:not(:first-child)');
  let hidden = 0, total = 0;
  rows.forEach(row => {
    total++;
    const text = row.textContent.toLowerCase();
    if (term && text.includes(term)) {
      row.classList.add('hidden-row');
      hidden++;
    } else {
      row.classList.remove('hidden-row');
    }
  });
  // Update badge counts per section and summary cards
  let totalVisible = 0, mismatchVisible = 0, t3OnlyVisible = 0, prodOnlyVisible = 0;
  document.querySelectorAll('.section').forEach(sec => {
    const tbl = sec.querySelector('table');
    if (!tbl) return;
    const visible = tbl.querySelectorAll('tr:not(:first-child):not(.hidden-row)').length;
    const badge = sec.querySelector('.badge');
    if (badge) badge.textContent = visible;
    const hdr = sec.querySelector('.section-header');
    if (!hdr) return;
    const txt = hdr.textContent.toLowerCase();
    totalVisible += visible;
    if (txt.includes('mismatch')) mismatchVisible += visible;
    else if (txt.includes('only in t3')) t3OnlyVisible += visible;
    else if (txt.includes('only in prod')) prodOnlyVisible += visible;
  });
  const cards = document.querySelectorAll('.card .num');
  if (cards.length >= 4) {
    cards[0].textContent = totalVisible;
    cards[1].textContent = mismatchVisible;
    cards[2].textContent = t3OnlyVisible;
    cards[3].textContent = prodOnlyVisible;
  }
  filterInfo.textContent = term ? `Hiding ${hidden} of ${total} rows` : '';
}

filterInput.addEventListener('input', applyFilter);

function togglePreset(btn) {
  const filter = btn.dataset.filter;
  if (btn.classList.contains('active')) {
    btn.classList.remove('active');
    filterInput.value = '';
  } else {
    document.querySelectorAll('.filter-btn[data-filter]').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    filterInput.value = filter;
  }
  applyFilter();
}

function clearFilter() {
  filterInput.value = '';
  document.querySelectorAll('.filter-btn[data-filter]').forEach(b => b.classList.remove('active'));
  applyFilter();
}
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    return output_path


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def compare_pair(t3_path: str, prod_path: str, base_name: str):
    """Run the full comparison pipeline for a single T3/PROD pair."""
    print(f"\n{'=' * 70}")
    print(f"  Comparing: {base_name}")
    print(f"{'=' * 70}")

    # --- Extract text ---
    print(f"\n[1/5] Extracting text from T3 PDF: {os.path.basename(t3_path)}")
    t3_text = extract_text_from_pdf(t3_path)
    print(f"      Extracted {len(t3_text)} characters, {len(t3_text.splitlines())} lines.")

    print(f"\n[2/5] Extracting text from PROD PDF: {os.path.basename(prod_path)}")
    prod_text = extract_text_from_pdf(prod_path)
    print(f"      Extracted {len(prod_text)} characters, {len(prod_text.splitlines())} lines.")

    # --- Quick equality check ---
    print("\n[3/5] Performing quick text equality check...")
    if t3_text.strip() == prod_text.strip():
        print("      *** Files are IDENTICAL. No differences found. ***")
    else:
        print("      Files differ. Proceeding with detailed analysis...")

    # --- Parse into sections ---
    print("\n[4/5] Parsing permission sections...")
    t3_sections = parse_sections(t3_text)
    prod_sections = parse_sections(prod_text)
    print(f"      T3 sections found:   {len(t3_sections)}")
    print(f"      PROD sections found: {len(prod_sections)}")

    # --- Compare ---
    print("\n[5/5] Comparing permissions...")
    differences = compare_permissions(t3_sections, prod_sections)
    raw_diff = compare_raw_lines(t3_text, prod_text)

    print(f"      Structured differences: {len(differences)}")
    print(f"      Raw lines only in T3:   {len(raw_diff['only_in_t3'])}")
    print(f"      Raw lines only in PROD: {len(raw_diff['only_in_prod'])}")

    # --- Generate Reports ---
    output_txt = os.path.join(SCRIPT_DIR, f"{base_name}_comparison_report.txt")
    output_html = os.path.join(SCRIPT_DIR, f"{base_name}_comparison_report.html")

    print(f"\nGenerating reports...")
    report = generate_report(differences, raw_diff, output_txt,
                             t3_path=t3_path, prod_path=prod_path,
                             t3_sections=t3_sections,
                             prod_sections=prod_sections)
    generate_html_report(differences, raw_diff, output_html,
                         t3_path=t3_path, prod_path=prod_path)

    print("\n" + report)
    print(f"\n[DONE] Reports saved:")
    print(f"  Text: {os.path.basename(output_txt)}")
    print(f"  HTML: {os.path.basename(output_html)}")
    return output_txt, output_html


def main():
    print("=" * 70)
    print("  PDF Permission Comparator: T3 vs PROD")
    print("=" * 70)

    pairs = find_pdf_pairs()
    print(f"\nFound {len(pairs)} T3/PROD pair(s) to compare.")

    reports = []
    for t3_path, prod_path, base_name in pairs:
        txt_path, html_path = compare_pair(t3_path, prod_path, base_name)
        reports.append((txt_path, html_path))

    print(f"\n{'=' * 70}")
    print(f"  All done! Generated {len(reports)} report(s):")
    for txt_path, html_path in reports:
        print(f"    - {os.path.basename(txt_path)}")
        print(f"    - {os.path.basename(html_path)}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    main()