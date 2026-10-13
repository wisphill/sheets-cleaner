"""Compare the Employees / Clients sheets of Raykay_transaction.xlsx against the
source exports in data/comparison and write a diff report (xlsx).

Records are matched by email first, then by any phone number (and by Uid for
employees). The report contains, for each entity:
  - Missing: rows in the source exports not found in Raykay_transaction
  - Extra:   rows in Raykay_transaction not found in any source export
  - Matched: matched pairs, with how they were matched and name differences

Usage (from repo root):
    python3 scripts/compare_raykay_transaction.py [--transaction PATH] [--out PATH]
"""

import argparse
import csv
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Set

import openpyxl
from openpyxl.styles import Font, PatternFill

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
COMPARISON_DIR = os.path.join(ROOT, "data", "comparison")

DEFAULT_TRANSACTION = os.path.join(ROOT, "Raykay_transaction.xlsx")
DEFAULT_EMPLOYEE_SOURCES = [
    os.path.join(COMPARISON_DIR, "employee-list Active 13.10.2026.xlsx"),
    os.path.join(COMPARISON_DIR, "employee-KP active 13.10.2026.xlsx"),
]
DEFAULT_CLIENT_SOURCE = os.path.join(COMPARISON_DIR, "client-list 13.10. 3.30pm.csv")

EMAIL_RE = re.compile(r"[\w.+'-]+@[\w-]+(?:\.[\w-]+)+")
PHONE_RE = re.compile(r"\+?\d[\d\s\-().]{6,}\d")
# Last 8 digits identify an AU number regardless of +61 / 0 / area code prefix.
PHONE_KEY_DIGITS = 8


@dataclass
class Record:
    source: str
    row: int
    id: str
    first_name: str
    last_name: str
    emails: Set[str] = field(default_factory=set)
    phones: Set[str] = field(default_factory=set)
    uid: str = ""
    raw_emails: str = ""
    raw_phones: str = ""

    @property
    def full_name(self) -> str:
        return f"{self.first_name} {self.last_name}".strip()

    @property
    def name_key(self) -> str:
        return normalize_name(self.full_name)


def to_str(val) -> str:
    return "" if val is None else str(val).strip()


def normalize_name(name: str) -> str:
    return re.sub(r"[^a-z]", "", name.lower())


def extract_emails(*values) -> Set[str]:
    emails = set()
    for val in values:
        emails.update(m.lower() for m in EMAIL_RE.findall(to_str(val)))
    return emails


def extract_phones(*values) -> Set[str]:
    phones = set()
    for val in values:
        for match in PHONE_RE.findall(to_str(val)):
            digits = re.sub(r"\D", "", match)
            if len(digits) >= PHONE_KEY_DIGITS:
                phones.add(digits[-PHONE_KEY_DIGITS:])
    return phones


def join_raw(*values) -> str:
    return " | ".join(to_str(v) for v in values if to_str(v) not in ("", "0"))


# ==============================================================================
# Loaders
# ==============================================================================
def read_sheet_dicts(path: str, sheet_name: Optional[str] = None, skip_rows: int = 0):
    """Yield (excel_row_number, {header: value}) for each non-empty data row."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    headers = [to_str(h) for h in next(rows)]
    for idx, values in enumerate(rows, start=2):
        if idx < 2 + skip_rows:
            continue
        if all(to_str(v) == "" for v in values):
            continue
        yield idx, dict(zip(headers, values))
    wb.close()


def load_transaction_employees(path: str) -> List[Record]:
    records = []
    phone_cols = ["MainPhoneNumber", "PersonalPhoneNumber", "OtherPhoneNumber"]
    # Row 2 of the import template holds column descriptions.
    for row, d in read_sheet_dicts(path, "Employees", skip_rows=1):
        if not to_str(d.get("Id*")):
            continue
        records.append(Record(
            source="Raykay_transaction/Employees",
            row=row,
            id=to_str(d.get("Id*")),
            first_name=to_str(d.get("FirstName*")),
            last_name=to_str(d.get("LastName*")),
            emails=extract_emails(d.get("EmailAddress*")),
            phones=extract_phones(*(d.get(c) for c in phone_cols)),
            uid=to_str(d.get("uid")),
            raw_emails=join_raw(d.get("EmailAddress*")),
            raw_phones=join_raw(*(d.get(c) for c in phone_cols)),
        ))
    return records


def load_transaction_clients(path: str) -> List[Record]:
    records = []
    phone_cols = ["MainPhoneNumber", "PersonalPhoneNumber", "OtherPhoneNumber"]
    for row, d in read_sheet_dicts(path, "Clients", skip_rows=1):
        if not to_str(d.get("Id*")):
            continue
        records.append(Record(
            source="Raykay_transaction/Clients",
            row=row,
            id=to_str(d.get("Id*")),
            first_name=to_str(d.get("FirstName*")),
            last_name=to_str(d.get("LastName*")),
            emails=extract_emails(d.get("EmailAddress")),
            phones=extract_phones(*(d.get(c) for c in phone_cols)),
            raw_emails=join_raw(d.get("EmailAddress")),
            raw_phones=join_raw(*(d.get(c) for c in phone_cols)),
        ))
    return records


def load_employee_source(path: str) -> List[Record]:
    records = []
    name = os.path.basename(path)
    for row, d in read_sheet_dicts(path):
        records.append(Record(
            source=name,
            row=row,
            id=to_str(d.get("Uid")),
            first_name=to_str(d.get("Name")),
            last_name=to_str(d.get("Last  Name")),
            emails=extract_emails(d.get("Email")),
            phones=extract_phones(d.get("Phone")),
            uid=to_str(d.get("Uid")),
            raw_emails=join_raw(d.get("Email")),
            raw_phones=join_raw(d.get("Phone")),
        ))
    return records


def load_client_source(path: str) -> List[Record]:
    records = []
    name = os.path.basename(path)
    phone_cols = ["phone_main", "phone_personal", "phone_other"]
    with open(path, encoding="cp1252", newline="") as f:
        for row, d in enumerate(csv.DictReader(f), start=2):
            records.append(Record(
                source=name,
                row=row,
                id=to_str(d.get("id")),
                first_name=to_str(d.get("first_name")),
                last_name=to_str(d.get("last_name")),
                emails=extract_emails(d.get("email")),
                phones=extract_phones(*(d.get(c) for c in phone_cols)),
                raw_emails=join_raw(d.get("email")),
                raw_phones=join_raw(*(d.get(c) for c in phone_cols)),
            ))
    return records


# ==============================================================================
# Matching
# ==============================================================================
@dataclass
class MatchResult:
    matched: List[tuple]  # (source_record, target_record, method)
    missing: List[Record]
    extra: List[Record]


def build_index(records: List[Record], key_fn) -> Dict[str, List[Record]]:
    index: Dict[str, List[Record]] = {}
    for rec in records:
        for key in key_fn(rec):
            index.setdefault(key, []).append(rec)
    return index


def compare(sources: List[Record], targets: List[Record]) -> MatchResult:
    strategies = [
        ("email", lambda r: r.emails),
        ("phone", lambda r: r.phones),
        ("uid", lambda r: {r.uid} if r.uid else set()),
    ]
    indexes = [(name, fn, build_index(targets, fn)) for name, fn in strategies]

    matched, missing = [], []
    matched_target_ids: Set[int] = set()

    for src in sources:
        found = None
        for method, key_fn, index in indexes:
            candidates = [t for key in key_fn(src) for t in index.get(key, [])]
            if not candidates:
                continue
            # Prefer a candidate with the same name, e.g. when a phone is shared.
            same_name = [t for t in candidates if t.name_key == src.name_key]
            unused = [t for t in (same_name or candidates) if id(t) not in matched_target_ids]
            found = (unused or same_name or candidates)[0]
            matched.append((src, found, method))
            matched_target_ids.add(id(found))
            break
        if found is None:
            missing.append(src)

    extra = [t for t in targets if id(t) not in matched_target_ids]
    return MatchResult(matched, missing, extra)


def name_hints(records: List[Record], others: List[Record]) -> Dict[int, str]:
    """For unmatched records, list records on the other side with the same name."""
    by_name = build_index(others, lambda r: {r.name_key} if r.name_key else set())
    hints = {}
    for rec in records:
        hits = by_name.get(rec.name_key, [])
        hints[id(rec)] = "; ".join(
            f"{h.id} ({h.raw_emails or '-'} / {h.raw_phones or '-'})" for h in hits
        )
    return hints


# ==============================================================================
# Report
# ==============================================================================
HEADER_FONT = Font(bold=True)
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")


def write_table(ws, headers: List[str], rows: List[list], highlight_col: Optional[int] = None):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    for r in rows:
        ws.append(r)
        if highlight_col is not None and r[highlight_col]:
            for cell in ws[ws.max_row]:
                cell.fill = WARN_FILL
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        width = max(len(to_str(c.value)) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 60)


def unmatched_rows(records: List[Record], hints: Dict[int, str]) -> List[list]:
    return [
        [r.source, r.row, r.id, r.first_name, r.last_name, r.raw_emails, r.raw_phones, hints[id(r)]]
        for r in records
    ]


def add_entity_sheets(wb, label: str, sources: List[Record], targets: List[Record], result: MatchResult):
    unmatched_headers = [
        "Source", "Row", "Id", "FirstName", "LastName", "Email", "Phones",
        "Same name on other side (possible match)",
    ]

    ws = wb.create_sheet(f"{label} Missing")
    write_table(ws, unmatched_headers,
                unmatched_rows(result.missing, name_hints(result.missing, targets)),
                highlight_col=7)

    ws = wb.create_sheet(f"{label} Extra")
    write_table(ws, unmatched_headers,
                unmatched_rows(result.extra, name_hints(result.extra, sources)),
                highlight_col=7)

    ws = wb.create_sheet(f"{label} Matched")
    rows = []
    for src, tgt, method in result.matched:
        name_diff = "YES" if src.name_key != tgt.name_key else ""
        rows.append([
            method, name_diff,
            src.source, src.id, src.full_name, src.raw_emails, src.raw_phones,
            tgt.id, tgt.full_name, tgt.raw_emails, tgt.raw_phones,
        ])
    write_table(ws, [
        "Matched by", "Name differs",
        "Source", "Source Id", "Source Name", "Source Email", "Source Phones",
        "Transaction Id", "Transaction Name", "Transaction Email", "Transaction Phones",
    ], rows, highlight_col=1)


def summary_rows(label: str, sources: List[Record], targets: List[Record], result: MatchResult) -> List[list]:
    by_method: Dict[str, int] = {}
    for _, _, method in result.matched:
        by_method[method] = by_method.get(method, 0) + 1
    name_diffs = sum(1 for s, t, _ in result.matched if s.name_key != t.name_key)
    return [
        [label, "Source rows", len(sources)],
        [label, "Transaction rows", len(targets)],
        [label, "Matched", len(result.matched)],
        *[[label, f"  matched by {m}", n] for m, n in by_method.items()],
        [label, "Matched but name differs", name_diffs],
        [label, "Missing from transaction", len(result.missing)],
        [label, "Extra in transaction (not in sources)", len(result.extra)],
    ]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--transaction", default=DEFAULT_TRANSACTION)
    parser.add_argument("--employees", nargs="+", default=DEFAULT_EMPLOYEE_SOURCES)
    parser.add_argument("--clients", default=DEFAULT_CLIENT_SOURCE)
    parser.add_argument("--out", default=os.path.join(
        COMPARISON_DIR, f"comparison_report_{datetime.now():%Y-%m-%d_%H%M%S}.xlsx"))
    args = parser.parse_args()

    emp_sources = [r for p in args.employees for r in load_employee_source(p)]
    emp_targets = load_transaction_employees(args.transaction)
    emp_result = compare(emp_sources, emp_targets)

    cli_sources = load_client_source(args.clients)
    cli_targets = load_transaction_clients(args.transaction)
    cli_result = compare(cli_sources, cli_targets)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    summary = (summary_rows("Employees", emp_sources, emp_targets, emp_result)
               + summary_rows("Clients", cli_sources, cli_targets, cli_result))
    write_table(ws, ["Entity", "Metric", "Count"], summary)

    add_entity_sheets(wb, "Employees", emp_sources, emp_targets, emp_result)
    add_entity_sheets(wb, "Clients", cli_sources, cli_targets, cli_result)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    wb.save(args.out)

    for entity, metric, count in summary:
        print(f"{entity:<10} {metric:<40} {count}")
    print(f"\n✅ Report written to {args.out}")


if __name__ == "__main__":
    main()
