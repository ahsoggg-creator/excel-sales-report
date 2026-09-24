import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

PRODUCTS = {
    "Ноутбук Lenovo IdeaPad": 54990,
    "Мышь Logitech M185": 1290,
    "Клавиатура Defender": 890,
    "Монитор Samsung 24\"": 13490,
    "Наушники JBL Tune 510": 3990,
    "Флешка Kingston 64GB": 690,
    "Роутер TP-Link Archer": 2790,
}
MANAGERS = ["Иванов", "Петрова", "Сидоров", "Кузнецова", "Смирнов"]

BRANCHES = {
    "Москва": {"qty_col": "Количество", "date_as_text": False},
    "Казань": {"qty_col": "Кол-во", "date_as_text": True},
    "Новосибирск": {"qty_col": "Кол-во, шт", "date_as_text": False},
    "Екатеринбург": {"qty_col": "Количество", "date_as_text": True},
}


def random_rows(count):
    start = date(2026, 1, 1)
    rows = []
    for _ in range(count):
        product, price = random.choice(list(PRODUCTS.items()))
        rows.append([
            start + timedelta(days=random.randint(0, 180)),
            product,
            random.randint(1, 10),
            price,
            random.choice(MANAGERS),
        ])
    rows.sort(key=lambda r: r[0])
    return rows


def spoil(rows, date_as_text):
    result = []
    for row in rows:
        row = list(row)
        if date_as_text:
            row[0] = row[0].strftime(random.choice(["%d.%m.%Y", "%Y-%m-%d"]))
        if random.random() < 0.15:
            row[1] = "  " + row[1] + " "
        if random.random() < 0.2:
            row[3] = f"{row[3]:,}".replace(",", " ") + " р."
        if random.random() < 0.1:
            row[2] = str(row[2])
        result.append(row)
        if random.random() < 0.08:
            result.append(list(row))
        if random.random() < 0.05:
            result.append([None] * 5)
    return result


def main():
    random.seed(42)
    folder = Path("input")
    folder.mkdir(exist_ok=True)

    for branch, settings in BRANCHES.items():
        wb = Workbook()
        ws = wb.active
        ws.title = "Продажи"
        ws.append([f"Отчет по продажам. Филиал: {branch}"])
        ws.append([])
        ws.append(["Дата", "Товар", settings["qty_col"], "Цена", "Менеджер"])

        rows = spoil(random_rows(random.randint(60, 90)), settings["date_as_text"])
        for row in rows:
            ws.append(row)

        ws.append([])
        ws.append(["Итого", None, None, None, None])
        wb.save(folder / f"{branch}.xlsx")
        print(f"{branch}: {len(rows)} строк")


if __name__ == "__main__":
    main()
