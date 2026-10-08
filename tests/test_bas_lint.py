"""Static checks on src/ReportAutomation.bas text -- VBA itself cannot run in CI."""
import os
import re
import sys
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

BAS_PATH = os.path.join(ROOT, "src", "ReportAutomation.bas")

SUB_FUNCTION_RE = re.compile(
    r"^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+(\w+)",
    re.IGNORECASE | re.MULTILINE,
)


def _read_bas():
    with open(BAS_PATH, "r", encoding="utf-8") as f:
        return f.read()


def test_bas_file_exists():
    assert os.path.isfile(BAS_PATH)


def test_no_continue_do_which_is_not_valid_vba():
    text = _read_bas()
    assert "Continue Do" not in text


def test_no_duplicate_sub_or_function_names():
    text = _read_bas()
    names = [name.lower() for name in SUB_FUNCTION_RE.findall(text)]
    assert names, "expected at least one Sub or Function definition"
    duplicates = [name for name, count in Counter(names).items() if count > 1]
    assert duplicates == []


def test_option_explicit_is_present():
    text = _read_bas()
    assert "Option Explicit" in text


def test_no_hardcoded_c_drive_paths():
    text = _read_bas()
    assert "C:\\" not in text


def _transfer_body():
    text = _read_bas()
    start = re.search(r"^\s*Sub\s+TransferToWordTemplate", text, re.I | re.M).start()
    return text[start : re.search(r"^\s*End Sub", text[start:], re.I | re.M).end() + start]


def test_fresh_document_is_created_inside_the_row_loop():
    body = _transfer_body()
    loop = re.search(r"^\s*For i = .*?^\s*Next i", body, re.I | re.M | re.S)
    assert loop, "expected a For i loop"
    assert "Documents.Add" in loop.group(0)
    assert "Documents.Add" not in body.replace(loop.group(0), "")


def test_missing_bookmarks_are_reported_not_silently_skipped():
    body = _transfer_body()
    assert ".Exists(" not in body
    assert "FillBookmark" in body
    assert "missing" in _read_bas().lower()
