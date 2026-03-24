from pathlib import Path
import pandas as pd


def get_engine(path: str) -> str:
    ext = Path(path).suffix.lower()
    if ext == ".xls":
        return "xlrd"
    if ext == ".xlsx":
        return "openpyxl"
    raise ValueError("Поддерживаются только .xls и .xlsx")


def inspect_excel(path: str):
    engine = get_engine(path)
    df = pd.read_excel(path, header=None, engine=engine)

    print(f"Файл: {path}")
    print(f"Engine: {engine}")
    print(f"Размер: {df.shape[0]} строк, {df.shape[1]} колонок")
    print("\n================ ПОЛНЫЙ DATAFRAME ================\n")
    print(df)

    print("\n================ ПОСТРОЧНЫЙ ВЫВОД ================\n")
    for row_idx in range(len(df)):
        row = df.iloc[row_idx].tolist()
        print(f"ROW {row_idx}:")
        for col_idx, value in enumerate(row):
            print(f"  col={col_idx} value={repr(value)}")
        print("-" * 50)


if __name__ == "__main__":
    path = "Счёт орг.стекло ИП Козин.xls"
    try:
        inspect_excel(path)
    except Exception as e:
        print(f"Ошибка при чтении файла: {e}")