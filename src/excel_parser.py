from pathlib import Path
import re

import pandas as pd


def normalize_text(value):
    if value is None:
        return ""

    text = str(value).replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", text)


class InvoiceParser:
    PRODUCT_HEADER_KEYWORDS = [
        "наименование товара",
        "наименование продукции",
        "наименование материальных ценностей",
        "товар",
        "наименование",
        "работ, услуг",
    ]
    QTY_HEADER_KEYWORDS = ["количество", "кол-во", "кол во", "qty"]
    UNIT_HEADER_KEYWORDS = ["ед. изм", "ед изм", "единица", "ед.", "шт", "шт."]

    def read_excel(self, path):
        ext = Path(path).suffix.lower()

        if ext == ".xls":
            return pd.read_excel(path, header=None, engine="xlrd")
        if ext == ".xlsx":
            return pd.read_excel(path, header=None, engine="openpyxl")

        raise ValueError("Поддерживаются только .xls и .xlsx")

    def find_header_row(self, df):
        for idx in range(len(df)):
            row_values = [normalize_text(x).lower() for x in df.iloc[idx].tolist()]
            row_text = " | ".join(row_values)

            has_name = any(keyword in row_text for keyword in self.PRODUCT_HEADER_KEYWORDS)
            has_qty = any(keyword in row_text for keyword in self.QTY_HEADER_KEYWORDS)

            if has_name and has_qty:
                return idx

        return None

    def find_column(self, headers, keywords):
        for index, header in enumerate(headers):
            value = normalize_text(header).lower()
            if any(keyword in value for keyword in keywords):
                return index
        return None

    def is_summary_row(self, text):
        text = text.lower()
        markers = [
            "итого",
            "всего к оплате",
            "в том числе",
            "сумма",
            "ндс",
            "всего наименований",
        ]
        return any(marker in text for marker in markers)

    def parse_items(self, path):
        df = self.read_excel(path)
        header_row = self.find_header_row(df)

        if header_row is None:
            raise ValueError(
                "Не удалось найти строку заголовков. Нужны колонки с наименованием и количеством."
            )

        headers = df.iloc[header_row].tolist()
        name_col = self.find_column(headers, self.PRODUCT_HEADER_KEYWORDS)
        qty_col = self.find_column(headers, self.QTY_HEADER_KEYWORDS)
        unit_col = self.find_column(headers, self.UNIT_HEADER_KEYWORDS)

        if name_col is None or qty_col is None:
            raise ValueError("Не удалось определить колонки с наименованием и количеством.")

        items = []

        for idx in range(header_row + 1, len(df)):
            row = df.iloc[idx].tolist()

            if name_col >= len(row):
                continue

            name = normalize_text(row[name_col])
            if not name or name.lower() == "nan":
                continue

            if self.is_summary_row(name):
                break

            qty = row[qty_col] if qty_col < len(row) else ""
            unit = row[unit_col] if unit_col is not None and unit_col < len(row) else ""

            qty_text = normalize_text(qty)
            unit_text = normalize_text(unit)

            if qty_text and unit_text:
                quantity = f"{qty_text} {unit_text}"
            elif qty_text:
                quantity = qty_text
            else:
                quantity = ""

            items.append(
                {
                    "name": name,
                    "quantity": quantity,
                }
            )

        if not items:
            raise ValueError("Товары не найдены. Проверь структуру счёта.")

        return items