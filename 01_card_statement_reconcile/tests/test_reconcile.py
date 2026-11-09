import csv
import re
import shutil
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import reconcile as rc  # noqa: E402

MONTH = "2026-10"
SEEDED_CODES = {
    "UNKNOWN_CARD", "CARD_INACTIVE", "DUPLICATE_LEDGER", "MISSING_ASSIGNEE", "UNKNOWN_FACILITY",
    "ALLOCATION_MISSING", "ALLOCATION_INVALID", "AMOUNT_UNREADABLE", "DATE_OUT_OF_MONTH", "TOTAL_MISMATCH",
}


@pytest.fixture(scope="session")
def data_dir():
    subprocess.run([sys.executable, str(ROOT / "generate_dummy_data.py")], check=True)
    return ROOT / "data"


def read_posting_csv(path):
    with path.open(encoding="cp932", newline="") as file:
        return list(csv.DictReader(file))


@pytest.mark.parametrize(
    "amount,rates",
    [(10000, ["33.3", "33.3", "33.4"]), (1, ["50", "50"]), (9999, ["12.5", "37.5", "50"]), (0, ["70", "30"])],
)
def test_allocate_always_sums_to_original(amount, rates):
    shares = [(f"F{i}", Decimal(r)) for i, r in enumerate(rates)]
    assert sum(value for _, value in rc.allocate(amount, shares)) == amount


def test_allocate_gives_remainder_to_largest_fraction_then_top_row():
    shares = [("A", Decimal("50")), ("B", Decimal("50"))]
    assert rc.allocate(101, shares) == [("A", 51), ("B", 50)]


def test_normalize_card_handles_fullwidth_and_hyphen():
    assert rc.normalize_card("９９００-0000 0000-0001") == "9900000000000001"


def test_parse_marks_unreadable_amount_without_guessing():
    text = "対象月: 2026年10月\n1 2026/10/01 9900-0000-0000-0001 架空SS 20.5L 4,2O0\n請求合計: ¥4,200"
    statement = rc.parse_statement_text(text)
    assert statement.rows[0].amount is None and statement.total == 4200


def test_parse_reports_row_it_cannot_understand():
    statement = rc.parse_statement_text("1 2026/10/01 カード番号なし 4,200")
    assert not statement.rows and len(statement.unparsed) == 1


def test_seeded_errors_are_all_detected_and_held(data_dir, tmp_path):
    result, path = rc.run(MONTH, data_dir / "input", tmp_path)
    assert {i.code for i in result.issues} == SEEDED_CODES
    assert path.name.endswith("_要確認.csv")
    assert result.unreadable_count == 1
    assert result.read_total == result.posted_total + result.held_total
    assert sum(int(r["金額"]) for r in read_posting_csv(path)) == result.posted_total


def test_after_corrections_total_matches_pdf_and_file_is_final(data_dir, tmp_path):
    draft_result, draft = rc.run(MONTH, data_dir / "input", tmp_path)
    assert draft.exists()
    result, path = rc.run(MONTH, data_dir / "input_fixed", tmp_path)
    assert result.is_final and result.held_total == 0
    assert result.posted_total == result.pdf_total
    assert sum(int(r["金額"]) for r in read_posting_csv(path)) == result.pdf_total
    assert path.name == f"振替_{MONTH}.csv"
    assert not draft.exists(), "確定後に要確認ファイルが残ってはいけない"


def test_error_run_removes_earlier_final_file(data_dir, tmp_path):
    _, final = rc.run(MONTH, data_dir / "input_fixed", tmp_path)
    rc.run(MONTH, data_dir / "input", tmp_path)
    assert not final.exists()


def test_rerun_is_repeatable(data_dir, tmp_path):
    _, first = rc.run(MONTH, data_dir / "input_fixed", tmp_path)
    content = first.read_bytes()
    _, second = rc.run(MONTH, data_dir / "input_fixed", tmp_path)
    assert second.read_bytes() == content


def test_card_reassigned_by_period_uses_current_row(data_dir, tmp_path):
    result, _ = rc.run(MONTH, data_dir / "input_fixed", tmp_path)
    row = next(r for r in result.allocations if r["カード"] == "****-0024")
    assert row["施設コード"] == "F002"


def test_outputs_never_contain_full_card_numbers(data_dir, tmp_path):
    rc.run(MONTH, data_dir / "input", tmp_path)
    folder = tmp_path / MONTH
    full_number = re.compile(r"\d{4}[- ]?\d{4}[- ]?\d{4}")
    for path in folder.glob("*.csv"):
        assert not full_number.search(path.read_bytes().decode("cp932", errors="ignore"))
        assert not full_number.search(path.read_bytes().decode("utf-8", errors="ignore"))
    book = load_workbook(folder / f"集計表_{MONTH}.xlsx")
    for sheet in book:
        for row in sheet.iter_rows(values_only=True):
            assert not any(full_number.search(str(cell)) for cell in row if cell is not None)


def test_wrong_month_stops_processing(data_dir, tmp_path):
    source = data_dir / "input"
    copy = tmp_path / "in"
    shutil.copytree(source, copy)
    (copy / "statements" / f"{MONTH}.pdf").rename(copy / "statements" / "2026-11.pdf")
    with pytest.raises(rc.InputFileError, match="対象月"):
        rc.run("2026-11", copy, tmp_path / "out")


def test_missing_ledger_column_gives_clear_message(data_dir, tmp_path):
    copy = tmp_path / "in"
    shutil.copytree(data_dir / "input", copy)
    book = load_workbook(copy / "ledger.xlsx")
    book["台帳"]["B1"] = "種類"
    book.save(copy / "ledger.xlsx")
    with pytest.raises(rc.InputFileError, match="区分"):
        rc.run(MONTH, copy, tmp_path / "out")
