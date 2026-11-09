"""ガソリンカード明細の照合・按分・振替CSV作成ツール（サンプル）。

利用明細PDF、カード管理台帳、按分表を読み、施設別の費用振替CSVを作る。
判断できないデータは推測で処理せず、エラー一覧に出して振替から除外する。

使い方:
    python reconcile.py --month 2026-10 --input data/input --output output
"""
from __future__ import annotations

import argparse
import calendar
import csv
import re
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from pypdf import PdfReader

ROW_START = re.compile(r"^\s*\d+\s+\d{4}/\d{2}/\d{2}\b")
ROW = re.compile(
    r"^\s*(?P<no>\d+)\s+(?P<date>\d{4}/\d{2}/\d{2})\s+"
    r"(?P<card>\d{4}-\d{4}-\d{4}-\d{4})\s+(?P<station>.+?)\s+"
    r"(?P<liters>[\d.]+)\s*L\s+(?P<amount>\S+)\s*$"
)
MONTH_LINE = re.compile(r"対象月\s*[:：]?\s*(\d{4})年\s*(\d{1,2})月")
TOTAL_LINE = re.compile(r"請求合計\s*[:：]?\s*[¥￥]?\s*([\d,]+)")

ERROR_GUIDE = {
    "UNKNOWN_CARD": "カード管理台帳に登録してください。",
    "CARD_INACTIVE": "台帳の有効期間、または利用の妥当性を確認してください。",
    "DUPLICATE_LEDGER": "同じ期間に有効な行を1行にしてください。",
    "MISSING_ASSIGNEE": "台帳の区分、施設コード、利用者を入力してください。",
    "UNKNOWN_FACILITY": "施設コードを修正するか、施設マスタに追加してください。",
    "ALLOCATION_MISSING": "按分表に利用者の按分先を追加してください。",
    "ALLOCATION_INVALID": "割合の合計を100にし、施設コードを確認してください。",
    "AMOUNT_UNREADABLE": "原本を確認し、corrections.csv に金額を入力してください。",
    "DATE_OUT_OF_MONTH": "原本を確認し、corrections.csv に利用日を入力してください。",
    "DUPLICATE_ROW": "同じ明細Noが複数あります。原本を確認してください。",
    "ROW_UNPARSED": "PDFの行を読み取れませんでした。原本を確認してください。",
    "TOTAL_MISMATCH": "明細の合計と請求合計が一致しません。他のエラーを解消して再実行してください。",
    "TOTAL_NOT_FOUND": "PDFから請求合計を読み取れませんでした。原本を確認してください。",
}


class InputFileError(Exception):
    """入力ファイルが処理を続けられない状態のときに送出する。"""


@dataclass(frozen=True)
class Row:
    no: int
    used_on: date
    card: str
    station: str
    amount: int | None
    raw_amount: str


@dataclass
class Statement:
    year: int | None = None
    month: int | None = None
    total: int | None = None
    rows: list[Row] = field(default_factory=list)
    unparsed: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class LedgerEntry:
    card: str
    kind: str
    facility: str
    user: str
    start: date | None
    end: date | None


@dataclass(frozen=True)
class Issue:
    code: str
    detail: str
    card: str = ""
    row_no: int | None = None
    amount: int | None = None


@dataclass
class Result:
    month: str
    pdf_total: int | None
    read_total: int
    posted_total: int
    held_total: int
    unreadable_count: int
    by_card: list[dict]
    by_facility: dict[str, int]
    allocations: list[dict]
    issues: list[Issue]

    @property
    def is_final(self) -> bool:
        return not self.issues


def normalize_card(value: object) -> str:
    """全角やハイフンの揺れを除き、カード番号を16桁の数字にする。"""
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"[\s-]", "", text)


def mask_card(card: str) -> str:
    """出力に全桁を残さないよう、下4桁だけを表示する。"""
    return f"****-{card[-4:]}" if len(card) >= 4 else card


def parse_statement_text(text: str) -> Statement:
    statement = Statement()
    for line in text.splitlines():
        if statement.year is None and (found := MONTH_LINE.search(line)):
            statement.year, statement.month = int(found[1]), int(found[2])
            continue
        if found := TOTAL_LINE.search(line):
            statement.total = int(found[1].replace(",", ""))
            continue
        if not ROW_START.match(line):
            continue
        found = ROW.match(line)
        if found is None:
            statement.unparsed.append(line.strip())
            continue
        raw = found["amount"].lstrip("¥￥")
        amount = int(raw.replace(",", "")) if re.fullmatch(r"\d[\d,]*", raw) else None
        statement.rows.append(
            Row(
                no=int(found["no"]),
                used_on=datetime.strptime(found["date"], "%Y/%m/%d").date(),
                card=normalize_card(found["card"]),
                station=found["station"],
                amount=amount,
                raw_amount=raw,
            )
        )
    return statement


def read_statement_pdf(path: Path) -> Statement:
    reader = PdfReader(path)
    # 既定の抽出では表の各セルが別の行になるため、レイアウトを保つモードで読む。
    pages = [page.extract_text(extraction_mode="layout") or "" for page in reader.pages]
    return parse_statement_text("\n".join(pages))


def _sheet_rows(path: Path, sheet: str, required: list[str]) -> list[dict]:
    if not path.exists():
        raise InputFileError(f"ファイルが見つかりません: {path}")
    workbook = load_workbook(path, data_only=True, read_only=True)
    if sheet not in workbook.sheetnames:
        raise InputFileError(f"{path.name} にシート「{sheet}」がありません。")
    rows = workbook[sheet].iter_rows(values_only=True)
    header = [str(cell).strip() if cell is not None else "" for cell in next(rows, ())]
    missing = [name for name in required if name not in header]
    if missing:
        raise InputFileError(f"{path.name} の「{sheet}」に列がありません: {', '.join(missing)}")
    records = [dict(zip(header, row)) for row in rows]
    workbook.close()
    return [r for r in records if any(v not in (None, "") for v in r.values())]


def _to_date(value: object) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip().replace("/", "-")
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError as error:
        raise InputFileError(f"日付として読めない値があります: {value}") from error


def _text(value: object) -> str:
    return "" if value is None else unicodedata.normalize("NFKC", str(value)).strip()


def load_ledger(path: Path) -> tuple[list[LedgerEntry], dict[str, str]]:
    rows = _sheet_rows(path, "台帳", ["カード番号", "区分", "施設コード", "利用者", "有効開始", "有効終了"])
    ledger = [
        LedgerEntry(
            card=normalize_card(r["カード番号"]),
            kind=_text(r["区分"]),
            facility=_text(r["施設コード"]),
            user=_text(r["利用者"]),
            start=_to_date(r["有効開始"]),
            end=_to_date(r["有効終了"]),
        )
        for r in rows
    ]
    master = _sheet_rows(path, "施設マスタ", ["施設コード", "施設名"])
    return ledger, {_text(r["施設コード"]): _text(r["施設名"]) for r in master}


def load_allocations(path: Path) -> dict[str, list[tuple[str, Decimal]]]:
    allocations: dict[str, list[tuple[str, Decimal]]] = defaultdict(list)
    for r in _sheet_rows(path, "按分", ["利用者", "施設コード", "割合"]):
        try:
            rate = Decimal(str(r["割合"]))
        except InvalidOperation as error:
            raise InputFileError(f"按分の割合が数値ではありません: {r['割合']}") from error
        allocations[_text(r["利用者"])].append((_text(r["施設コード"]), rate))
    return dict(allocations)


def load_corrections(path: Path) -> dict[int, dict]:
    """原本を人が確認した修正値を読む。ファイルは任意。"""
    if not path.exists():
        return {}
    corrections = {}
    with path.open(encoding="utf-8-sig", newline="") as file:
        for r in csv.DictReader(file):
            try:
                no = int(r["明細No"])
                amount = int(r["金額"].replace(",", "")) if r.get("金額") else None
                used_on = _to_date(r.get("利用日") or None)
            except (KeyError, ValueError) as error:
                raise InputFileError(f"corrections.csv の値を確認してください: {r}") from error
            corrections[no] = {"amount": amount, "used_on": used_on}
    return corrections


def apply_corrections(rows: list[Row], corrections: dict[int, dict]) -> list[Row]:
    fixed = []
    for row in rows:
        change = corrections.get(row.no)
        if change is None:
            fixed.append(row)
            continue
        amount = row.amount if change["amount"] is None else change["amount"]
        used_on = change["used_on"] or row.used_on
        fixed.append(Row(row.no, used_on, row.card, row.station, amount, row.raw_amount))
    return fixed


def allocate(amount: int, shares: list[tuple[str, Decimal]]) -> list[tuple[str, int]]:
    """最大剰余法で配分する。配分の合計は必ず元の金額と一致する。

    同じ大きさの端数が並んだときは、按分表で上にある行を優先する。
    """
    total_rate = sum(rate for _, rate in shares)
    exact = [amount * rate / total_rate for _, rate in shares]
    floors = [int(value) for value in exact]
    remainder = amount - sum(floors)
    order = sorted(range(len(shares)), key=lambda i: (-(exact[i] - floors[i]), i))
    for index in order[:remainder]:
        floors[index] += 1
    return [(shares[i][0], floors[i]) for i in range(len(shares))]


def _active_entries(entries: list[LedgerEntry], first: date, last: date) -> list[LedgerEntry]:
    return [
        e for e in entries
        if (e.start is None or e.start <= last) and (e.end is None or e.end >= first)
    ]


def reconcile(
    month: str,
    statement: Statement,
    ledger: list[LedgerEntry],
    facilities: dict[str, str],
    allocations: dict[str, list[tuple[str, Decimal]]],
) -> Result:
    year, mon = map(int, month.split("-"))
    if (statement.year, statement.month) != (year, mon):
        raise InputFileError(
            f"PDFの対象月（{statement.year}年{statement.month}月）と指定した月（{month}）が違います。"
        )
    first, last = date(year, mon, 1), date(year, mon, calendar.monthrange(year, mon)[1])

    issues: list[Issue] = [Issue("ROW_UNPARSED", f"読取不可の行: {line}") for line in statement.unparsed]
    seen: set[int] = set()
    card_totals: dict[str, int] = defaultdict(int)
    card_counts: dict[str, int] = defaultdict(int)
    unreadable = 0
    read_total = 0
    date_held_total = 0

    for row in statement.rows:
        if row.no in seen:
            issues.append(Issue("DUPLICATE_ROW", "明細Noが重複しています。", row.card, row.no, row.amount))
            continue
        seen.add(row.no)
        if row.amount is None:
            unreadable += 1
            issues.append(
                Issue("AMOUNT_UNREADABLE", f"金額を数値にできません（読取値: {row.raw_amount}）", row.card, row.no)
            )
            continue
        read_total += row.amount
        if not first <= row.used_on <= last:
            date_held_total += row.amount
            issues.append(
                Issue("DATE_OUT_OF_MONTH", f"利用日が対象月外です（{row.used_on}）", row.card, row.no, row.amount)
            )
            continue
        card_totals[row.card] += row.amount
        card_counts[row.card] += 1

    if statement.total is None:
        issues.append(Issue("TOTAL_NOT_FOUND", "請求合計の行がありません。"))
    elif statement.total != read_total:
        issues.append(
            Issue(
                "TOTAL_MISMATCH",
                f"請求合計 {statement.total:,}円 / 読取合計 {read_total:,}円"
                f"（差額 {statement.total - read_total:,}円、金額読取不可 {unreadable}行）",
            )
        )

    entries_by_card: dict[str, list[LedgerEntry]] = defaultdict(list)
    for entry in ledger:
        entries_by_card[entry.card].append(entry)

    by_card: list[dict] = []
    by_facility: dict[str, int] = defaultdict(int)
    detail: list[dict] = []
    held_total = date_held_total

    for card in sorted(card_totals):
        amount = card_totals[card]
        issue, targets, entry = _resolve_card(card, amount, entries_by_card, first, last, facilities, allocations)
        status = "保留" if issue else "振替済み"
        by_card.append(
            {
                "カード": mask_card(card),
                "件数": card_counts[card],
                "金額": amount,
                "利用者": entry.user if entry else "",
                "区分": entry.kind if entry else "",
                "状態": status,
            }
        )
        if issue:
            issues.append(issue)
            held_total += amount
            continue
        for code, value in targets:
            by_facility[code] += value
            detail.append(
                {"カード": mask_card(card), "利用者": entry.user, "施設コード": code, "施設名": facilities[code], "金額": value}
            )

    posted_total = sum(by_facility.values())
    return Result(
        month, statement.total, read_total, posted_total, held_total, unreadable,
        by_card, dict(sorted(by_facility.items())), detail, issues,
    )


def _resolve_card(card, amount, entries_by_card, first, last, facilities, allocations):
    """カード1枚の振替先を決める。戻り値は (問題, 振替先, 台帳行)。"""
    entries = entries_by_card.get(card, [])
    if not entries:
        return Issue("UNKNOWN_CARD", "台帳にないカードです。", card, amount=amount), [], None
    active = _active_entries(entries, first, last)
    if not active:
        return Issue("CARD_INACTIVE", "対象月に有効な台帳行がありません。", card, amount=amount), [], None
    if len(active) > 1:
        return Issue("DUPLICATE_LEDGER", "対象月に有効な台帳行が複数あります。", card, amount=amount), [], None

    entry = active[0]
    if entry.kind == "施設":
        if not entry.facility or not entry.user:
            return Issue("MISSING_ASSIGNEE", "施設コードまたは利用者が空欄です。", card, amount=amount), [], entry
        if entry.facility not in facilities:
            return Issue("UNKNOWN_FACILITY", f"施設コード {entry.facility} はマスタにありません。", card, amount=amount), [], entry
        return None, [(entry.facility, amount)], entry
    if entry.kind != "按分" or not entry.user:
        return Issue("MISSING_ASSIGNEE", "区分または利用者が正しく入力されていません。", card, amount=amount), [], entry

    shares = allocations.get(entry.user)
    if not shares:
        return Issue("ALLOCATION_MISSING", f"{entry.user} の按分先がありません。", card, amount=amount), [], entry
    total_rate = sum(rate for _, rate in shares)
    if total_rate != 100 or any(rate <= 0 for _, rate in shares):
        return Issue("ALLOCATION_INVALID", f"{entry.user} の割合の合計が {total_rate} です。", card, amount=amount), [], entry
    unknown = [code for code, _ in shares if code not in facilities]
    if unknown:
        return Issue("ALLOCATION_INVALID", f"按分先の施設コードが不明です: {', '.join(unknown)}", card, amount=amount), [], entry
    return None, allocate(amount, shares), entry


def write_outputs(result: Result, facilities: dict[str, str], output_dir: Path, account: str) -> Path:
    """結果を書き出し、振替CSVのパスを返す。要確認と確定の取り違えを防ぐ。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    final_csv = output_dir / f"振替_{result.month}.csv"
    draft_csv = output_dir / f"振替_{result.month}_要確認.csv"
    target, stale = (final_csv, draft_csv) if result.is_final else (draft_csv, final_csv)
    stale.unlink(missing_ok=True)

    year, mon = map(int, result.month.split("-"))
    with target.open("w", encoding="cp932", newline="") as file:
        writer = csv.writer(file, lineterminator="\r\n")
        writer.writerow(["対象月", "施設コード", "施設名", "勘定科目", "金額", "摘要"])
        for code, amount in result.by_facility.items():
            writer.writerow(
                [result.month, code, facilities[code], account, amount, f"{year}年{mon}月 ガソリンカード利用分"]
            )

    with (output_dir / f"エラー一覧_{result.month}.csv").open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, lineterminator="\r\n")
        writer.writerow(["種別", "明細No", "カード", "金額", "内容", "対応方法"])
        for issue in result.issues:
            writer.writerow(
                [
                    issue.code,
                    "" if issue.row_no is None else issue.row_no,
                    mask_card(issue.card) if issue.card else "",
                    "" if issue.amount is None else issue.amount,
                    issue.detail,
                    ERROR_GUIDE[issue.code],
                ]
            )

    workbook = Workbook()
    summary = workbook.active
    summary.title = "検算"
    for line in [
        ["項目", "金額・件数"],
        ["PDFの請求合計", result.pdf_total],
        ["読み取れた明細の合計", result.read_total],
        ["振替した合計", result.posted_total],
        ["保留した合計", result.held_total],
        ["金額を読み取れなかった行数", result.unreadable_count],
        ["エラー件数", len(result.issues)],
        ["状態", "確定" if result.is_final else "要確認"],
    ]:
        summary.append(line)
    _add_table(workbook, "カード別", result.by_card)
    _add_table(
        workbook, "施設別",
        [{"施設コード": c, "施設名": facilities[c], "金額": a} for c, a in result.by_facility.items()],
    )
    _add_table(workbook, "按分明細", result.allocations)
    workbook.save(output_dir / f"集計表_{result.month}.xlsx")
    return target


def _add_table(workbook: Workbook, title: str, rows: list[dict]) -> None:
    sheet = workbook.create_sheet(title)
    if rows:
        sheet.append(list(rows[0]))
        for row in rows:
            sheet.append(list(row.values()))


def run(month: str, input_dir: Path, output_root: Path, account: str = "車両費") -> tuple[Result, Path]:
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", month):
        raise InputFileError("--month は 2026-10 の形式で指定してください。")
    statement = read_statement_pdf_for(month, input_dir)
    statement.rows = apply_corrections(statement.rows, load_corrections(input_dir / "corrections.csv"))
    ledger, facilities = load_ledger(input_dir / "ledger.xlsx")
    allocations = load_allocations(input_dir / "allocation.xlsx")
    result = reconcile(month, statement, ledger, facilities, allocations)
    return result, write_outputs(result, facilities, output_root / month, account)


def read_statement_pdf_for(month: str, input_dir: Path) -> Statement:
    path = input_dir / "statements" / f"{month}.pdf"
    if not path.exists():
        raise InputFileError(f"ファイルが見つかりません: {path}")
    return read_statement_pdf(path)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="カード明細の照合・按分・振替CSV作成")
    parser.add_argument("--month", required=True, help="対象月（例: 2026-10）")
    parser.add_argument("--input", type=Path, default=Path("data/input"))
    parser.add_argument("--output", type=Path, default=Path("output"))
    parser.add_argument("--account", default="車両費", help="振替CSVの勘定科目")
    args = parser.parse_args(argv)

    try:
        result, csv_path = run(args.month, args.input, args.output, args.account)
    except InputFileError as error:
        print(f"処理を中止しました: {error}", file=sys.stderr)
        return 2

    print(f"対象月: {result.month}")
    print(f"PDFの請求合計: {result.pdf_total:,}円" if result.pdf_total is not None else "PDFの請求合計: 読取不可")
    print(f"振替: {result.posted_total:,}円 / 保留: {result.held_total:,}円 / 金額読取不可: {result.unreadable_count}行")
    print(f"エラー: {len(result.issues)}件")
    print(f"出力: {csv_path}")
    return 0 if result.is_final else 1


if __name__ == "__main__":
    sys.exit(main())
