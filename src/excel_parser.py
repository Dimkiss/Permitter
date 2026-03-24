from pathlib import Path

import pandas as pd


class InvoiceParser:
    def __init__(self):
        self.PRODUCT_HEADER_KEYWORDS = [
            "наименование товара, работ, услуг",
            "наименование товара",
            "товар",
            "услуг",
        ]
        self.QTY_HEADER_KEYWORDS = ["количество", "qty"]
        self.UNIT_HEADER_KEYWORDS = ["ед. изм.", "ед.изм.", "ед изм", "unit"]
        self.PRICE_HEADER_KEYWORDS = ["цена", "price"]
        self.TOTAL_HEADER_KEYWORDS = ["сумма", "total"]

    def read_excel(self, path):
        ext = Path(path).suffix.lower()
        if ext == ".xls":
            return pd.read_excel(path, header=None, engine="xlrd")
        if ext == ".xlsx":
            return pd.read_excel(path, header=None, engine="openpyxl")
        raise ValueError("Поддерживаются только .xls и .xlsx")

    def _normalize(self, value):
        return str(value).replace("\xa0", " ").strip().lower()

    def _contains_keyword(self, row_values, keywords):
        return any(any(keyword in cell for cell in row_values) for keyword in keywords)

    def find_header_row(self, df):
        for idx, row in df.iterrows():
            row_values = [self._normalize(val) for val in row if pd.notna(val)]
            if self._contains_keyword(row_values, self.PRODUCT_HEADER_KEYWORDS) and self._contains_keyword(
                row_values, self.QTY_HEADER_KEYWORDS
            ):
                return idx
        return None

    def _find_col_idx(self, header_row, keywords):
        for idx, value in enumerate(header_row):
            if pd.isna(value):
                continue
            cell = self._normalize(value)
            if any(keyword in cell for keyword in keywords):
                return idx
        return None

    def parse_invoice(self, path):
        df = self.read_excel(path)
        header_row = self.find_header_row(df)
        if header_row is None:
            raise ValueError("Не удалось найти строку заголовков.")

        header = df.iloc[header_row]
        name_idx = self._find_col_idx(header, self.PRODUCT_HEADER_KEYWORDS)
        qty_idx = self._find_col_idx(header, self.QTY_HEADER_KEYWORDS)
        unit_idx = self._find_col_idx(header, self.UNIT_HEADER_KEYWORDS)
        price_idx = self._find_col_idx(header, self.PRICE_HEADER_KEYWORDS)
        total_idx = self._find_col_idx(header, self.TOTAL_HEADER_KEYWORDS)

        required_columns = [name_idx, qty_idx, unit_idx, price_idx, total_idx]
        if any(idx is None for idx in required_columns):
            raise ValueError("Не удалось определить все нужные колонки в строке заголовков.")

        items = []
        for row_idx in range(header_row + 1, len(df)):
            row = df.iloc[row_idx]
            name = row[name_idx]
            qty = row[qty_idx]
            unit = row[unit_idx]
            price = row[price_idx]
            total = row[total_idx]

            if pd.isna(name):
                continue

            name_text = str(name).strip()
            if not name_text:
                continue
            if name_text.lower().startswith("итого"):
                break

            items.append(
                {
                    "name": name_text,
                    "qty": qty,
                    "unit": unit,
                    "price": price,
                    "total": total,
                }
            )

        return items


if __name__ == "__main__":
    parser = InvoiceParser()
    path = "../Счёт орг.стекло ИП Козин.xls"
    try:
        items = parser.parse_invoice(path)
        print(f"Найдено позиций: {len(items)}")
        for idx, item in enumerate(items, start=1):
            print(f"{idx}. {item['name']} | {item['qty']} | {item['unit']} | {item['price']} | {item['total']}")
    except Exception as e:
        print(f"Ошибка: {e}")
