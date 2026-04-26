import pandas as pd
import re
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
RAW_DIR = BASE_DIR / "data" / "raw"

files_config = [
    {"name": "1.xlsx", "sheets": ["Лист1"]},
    {"name": "2.xls",  "sheets": ["Лист1"]},
    {"name": "3.xlsx", "sheets": ["Лист1", "Лист2"]},
]

all_products = []

for file_info in files_config:
    file_path = RAW_DIR / file_info["name"]
    if not file_path.exists():
        continue
    engine = 'xlrd' if file_info["name"].endswith('.xls') else 'openpyxl'
    for sheet in file_info["sheets"]:
        try:
            header_row = 1 if file_info["name"] == "1.xlsx" else 0
            df = pd.read_excel(file_path, sheet_name=sheet, engine=engine, header=header_row)
            col = next((c for c in df.columns if isinstance(c, str) and 'товар' in c.lower()), None)
            if col:
                all_products.extend(df[col].dropna().tolist())
        except:
            continue

# Витягуємо останній токен після останньої цифри/ваги
candidates = {}
for name in all_products:
    name = str(name).strip()
    # Беремо все після останнього числа з одиницею виміру
    match = re.search(r'\d+[.,]?\d*\s?(?:г|кг|мл|л)\S*\s+(.*)', name)
    if match:
        tail = match.group(1).strip()
    else:
        # Беремо останнє слово або кілька
        parts = name.split()
        tail = ' '.join(parts[-2:]) if len(parts) >= 2 else parts[-1] if parts else ''
    
    if tail and len(tail) >= 2:
        candidates[tail] = candidates.get(tail, 0) + 1

# Сортуємо по частоті
sorted_candidates = sorted(candidates.items(), key=lambda x: -x[1])

# Виводимо топ кандидатів
print("\n🏷️  КАНДИДАТИ НА БРЕНДИ (відсортовано по частоті):")
print("=" * 50)
for brand, count in sorted_candidates[:100]:
    print(f"  {count:3d}x  '{brand}',")