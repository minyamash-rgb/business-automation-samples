"""動作確認用の架空データを作る。実在のカード、施設、人物は含まない。

data/input       : エラーをわざと含むデータ
data/input_fixed : エラーを人が修正した後のデータ
"""
from __future__ import annotations

import csv
import random
from pathlib import Path

from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

BASE = Path(__file__).parent / "data"
FONT_PATH = r"C:\Windows\Fonts\msgothic.ttc"
FACILITIES = {f"F00{i}": f"サンプル第{i}事業所" for i in range(1, 7)}
STATIONS = ["架空SS中央店", "架空SS北店", "架空SS南店", "架空SS東店"]
ROW_COUNT = 200
UNREADABLE_NO = 17
OUT_OF_MONTH_NO = 42


def card(n: int) -> str:
    return f"9900-0000-0000-{n:04d}"


def base_ledger() -> list[list]:
    """(カード, 区分, 施設コード, 利用者, 有効開始, 有効終了)"""
    rows = []
    for n in range(1, 21):
        rows.append([card(n), "施設", f"F00{(n - 1) % 6 + 1}", f"利用者{n:02d}", "2026-04-01", None])
    rows += [
        [card(21), "施設", "F099", "利用者21", "2026-04-01", None],
        [card(22), "施設", None, None, "2026-04-01", None],
        [card(23), "施設", "F001", "利用者23", "2026-04-01", "2026-09-30"],
        [card(24), "施設", "F001", "利用者24", "2026-04-01", "2026-09-30"],
        [card(24), "施設", "F002", "利用者24", "2026-10-01", None],
        [card(25), "按分", None, "管理者01", "2026-04-01", None],
        [card(26), "按分", None, "管理者02", "2026-04-01", None],
        [card(27), "按分", None, "管理者03", "2026-04-01", None],
        [card(28), "按分", None, "管理者04", "2026-04-01", None],
        [card(29), "施設", "F003", "利用者29", "2026-04-01", None],
        [card(29), "施設", "F004", "利用者29", "2026-04-01", None],
    ]
    return rows


def fixed_ledger() -> list[list]:
    rows = [r for r in base_ledger() if not (r[0] == card(29) and r[2] == "F004")]
    for row in rows:
        if row[0] == card(21):
            row[2] = "F001"
        elif row[0] == card(22):
            row[2], row[3] = "F005", "利用者22"
        elif row[0] == card(23):
            row[5] = None
    rows.append([card(99), "施設", "F006", "利用者99", "2026-10-01", None])
    return rows


BASE_ALLOCATION = [
    ["管理者01", "F001", 33.3],
    ["管理者01", "F002", 33.3],
    ["管理者01", "F003", 33.4],
    ["管理者02", "F004", 50],
    ["管理者02", "F005", 50],
    ["管理者03", "F005", 40],
    ["管理者03", "F006", 50],
]
FIXED_ALLOCATION = [
    *BASE_ALLOCATION[:5],
    ["管理者03", "F005", 50],
    ["管理者03", "F006", 50],
    ["管理者04", "F001", 25],
    ["管理者04", "F006", 75],
]


def build_rows() -> list[dict]:
    rng = random.Random(20261110)
    cards = [*range(1, 30), 99]
    rows = []
    for no in range(1, ROW_COUNT + 1):
        # 全カードが1回は出るよう、最初に一巡させる。
        number = cards[no - 1] if no <= len(cards) else rng.choice(cards)
        rows.append(
            {
                "no": no,
                "date": f"2026/10/{rng.randint(1, 31):02d}",
                "card": card(number),
                "station": rng.choice(STATIONS),
                "liters": round(rng.uniform(10, 55), 1),
                "amount": rng.randrange(1500, 9000, 10),
            }
        )
    rows[OUT_OF_MONTH_NO - 1]["date"] = "2026/09/30"
    rows[UNREADABLE_NO - 1]["amount"] = 4200  # PDF上は「4,2O0」と印字される行
    return rows


def write_pdf(path: Path, rows: list[dict]) -> None:
    pdfmetrics.registerFont(TTFont("Gothic", FONT_PATH, subfontIndex=0))
    path.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(path), pagesize=A4)
    columns = [40, 80, 160, 310, 420, 480]

    def header() -> float:
        pdf.setFont("Gothic", 13)
        pdf.drawString(40, 800, "ガソリンカード利用明細（架空データ）")
        pdf.setFont("Gothic", 10)
        pdf.drawString(40, 780, "対象月: 2026年10月")
        for x, title in zip(columns, ["No", "利用日", "カード番号", "給油所", "数量", "金額"]):
            pdf.drawString(x, 755, title)
        return 735.0

    y = header()
    for row in rows:
        # OCRの誤読を想定し、この行だけ数字の0を英字のOにする。
        amount = "4,2O0" if row["no"] == UNREADABLE_NO else f"{row['amount']:,}"
        values = [row["no"], row["date"], row["card"], row["station"], f"{row['liters']}L", amount]
        for x, value in zip(columns, values):
            pdf.drawString(x, y, str(value))
        y -= 16
        if y < 60:
            pdf.showPage()
            y = header()
    pdf.setFont("Gothic", 11)
    pdf.drawString(40, max(y - 10, 40), f"請求合計: ¥{sum(r['amount'] for r in rows):,}")
    pdf.save()


def write_ledger(path: Path, rows: list[list]) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "台帳"
    sheet.append(["カード番号", "区分", "施設コード", "利用者", "有効開始", "有効終了"])
    for row in rows:
        sheet.append(row)
    master = book.create_sheet("施設マスタ")
    master.append(["施設コード", "施設名"])
    for item in FACILITIES.items():
        master.append(list(item))
    book.save(path)


def write_allocation(path: Path, rows: list[list]) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "按分"
    sheet.append(["利用者", "施設コード", "割合"])
    for row in rows:
        sheet.append(row)
    book.save(path)


def main() -> None:
    rows = build_rows()
    for name, ledger, allocation in [
        ("input", base_ledger(), BASE_ALLOCATION),
        ("input_fixed", fixed_ledger(), FIXED_ALLOCATION),
    ]:
        folder = BASE / name
        folder.mkdir(parents=True, exist_ok=True)
        write_pdf(folder / "statements" / "2026-10.pdf", rows)
        write_ledger(folder / "ledger.xlsx", ledger)
        write_allocation(folder / "allocation.xlsx", allocation)

    with (BASE / "input_fixed" / "corrections.csv").open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["明細No", "金額", "利用日", "理由"])
        writer.writerow([UNREADABLE_NO, rows[UNREADABLE_NO - 1]["amount"], "", "原本で金額を確認"])
        writer.writerow([OUT_OF_MONTH_NO, "", "2026-10-01", "原本で利用日を確認"])
    print(f"架空データを作成しました: {BASE}")


if __name__ == "__main__":
    main()
