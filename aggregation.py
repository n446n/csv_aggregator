import csv
from pathlib import Path
from openpyxl import Workbook
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table

pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))

def read_rows(csv_file):
    try:
        with open(csv_file, encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    except UnicodeDecodeError:
        with open(csv_file, encoding="cp932") as f:
            rows = list(csv.DictReader(f))
    return rows

def sum_by(rows, key):
    totals = {}
    for i in rows:
        buy = i[key]
        totals[buy] = totals.get(buy, 0) + int(i["単価"]) * int(i["数量"])
    return totals

def get_total(pair):
    return pair[1]

def get_desc(rows, name):
    totals = sum_by(rows, name)
    ranking = sorted(totals.items(), key=get_total, reverse=True)
    return ranking

def save_csv(ranking, filename):
    with open(filename, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["名前", "合計"])
        for name, total in ranking:
            writer.writerow([name,total])

def save_xlsx(ranking, filename):
    wb = Workbook()
    ws = wb.active
    ws.append(["名前", "合計"])
    for name, total in ranking:
        ws.append([name,total])
    wb.save(filename)

def save_pdf(ranking, filename):
    data = [["名前", "合計"]]
    for name, total in ranking:
        data.append([name, total])
    table = Table(data)
    table.setStyle([("FONTNAME",(0, 0), (-1, -1), "HeiseiKakuGo-W5"), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey)])
    doc = SimpleDocTemplate(str(filename), pagesize=A4)
    doc.build([table])

def aggregation(name):
    for p in Path("集計したいcsv").glob("*.csv"):
        try:
            rows = read_rows(p)
            ranking = get_desc(rows, name)
            csv_file = Path("集計後/csv") / (name + "_" + p.name)
            save_csv(ranking, csv_file)
            print(p, "の集計を", csv_file, "に出力しました")
            xlsx_file = Path("集計後/xlsx") / (name + "_" + p.stem + ".xlsx")
            save_xlsx(ranking, xlsx_file)
            print(p, "の集計を", xlsx_file, "に出力しました")
            pdf_file = Path("集計後/pdf") / (name + "_" + p.stem + ".pdf")
            save_pdf(ranking, pdf_file)
            print(p, "の集計を", pdf_file, "に出力しました")
        except KeyError:
            print(p, "は列名が一致しなかったためスキップしました")
        except ValueError:
            print(p, "は数値が不正だったためスキップしました")

Path("集計後/csv").mkdir(parents=True, exist_ok=True)
Path("集計後/xlsx").mkdir(parents=True, exist_ok=True)
Path("集計後/pdf").mkdir(parents=True, exist_ok=True)

aggregation("店舗")
aggregation("商品")
input("Enterキーで終了...")