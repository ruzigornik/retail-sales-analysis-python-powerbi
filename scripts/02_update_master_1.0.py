import pandas as pd
from pathlib import Path
from datetime import datetime

# ==========================================
# 1. НАЛАШТУВАННЯ ШЛЯХІВ
# ==========================================
BASE_DIR   = Path(__file__).parent.parent
PROC_DIR   = BASE_DIR / "data" / "processed"
REF_DIR    = BASE_DIR / "data" / "reference"
MASTER_FILE = REF_DIR / "directory_master.xlsx"

if not REF_DIR.exists():
    REF_DIR.mkdir(parents=True, exist_ok=True)

# Беремо останній файл з categorized_products_unique
input_files = sorted(PROC_DIR.glob("categorized_products_unique_*.xlsx"))
if not input_files:
    print(" Файл categorized_products_unique не знайдено.")
    exit()

IN_FILE = input_files[-1]
print(f" Джерело: {IN_FILE.name}")

# ==========================================
# 2. ЧИТАЄМО АВТОМАТИЧНО ОБРОБЛЕНИЙ ФАЙЛ
# ==========================================
auto_df = pd.read_excel(IN_FILE)

# Залишаємо тільки потрібні колонки для довідника
cols = ['товар', 'Категорія', 'Підкатегорія', 'ТМ', 'Смак', 'Вага']
auto_df = auto_df[cols].copy()

# ==========================================
# 3. СТВОРЮЄМО АБО ОНОВЛЮЄМО MASTER
# ==========================================
if not MASTER_FILE.exists():
    # Перший запуск — створюємо master з поточних даних
    master_df = auto_df.copy()
    print(f" Master файл створено вперше: {len(master_df)} товарів")
else:
    # Master вже існує — оновлюємо тільки порожні клітинки
    master_df = pd.read_excel(MASTER_FILE)
    print(f" Master файл знайдено: {len(master_df)} товарів")

    # Додаємо нові товари яких ще немає в master
    new_items = auto_df[~auto_df['товар'].isin(master_df['товар'])]
    if len(new_items) > 0:
        master_df = pd.concat([master_df, new_items], ignore_index=True)
        print(f" Додано нових товарів: {len(new_items)}")

    # Оновлюємо порожні клітинки з auto_df
    master_df = master_df.set_index('товар')
    auto_df   = auto_df.set_index('товар')

    for col in ['Категорія', 'Підкатегорія', 'ТМ', 'Смак', 'Вага']:
        # Заповнюємо тільки де в master порожньо або NaN
        mask = master_df[col].isna() | (master_df[col] == '')
        master_df.loc[mask, col] = auto_df.loc[
            auto_df.index.isin(master_df[mask].index), col
        ]

    master_df = master_df.reset_index()
    print(f" Порожні клітинки оновлено з автоматичних даних")

# ==========================================
# 4. ЗБЕРІГАЄМО MASTER
# ==========================================
master_df.to_excel(MASTER_FILE, index=False)

# ==========================================
# 5. ЗВІТ ПО MASTER ФАЙЛУ
# ==========================================
timestamp = datetime.now().strftime("%Y%m%d_%H%M")

stats = []
for cat in sorted(master_df['Категорія'].dropna().unique()):
    for subcat in sorted(
        master_df[master_df['Категорія'] == cat]['Підкатегорія'].dropna().unique()
    ):
        subset = master_df[
            (master_df['Категорія'] == cat) &
            (master_df['Підкатегорія'] == subcat)
        ]
        total        = len(subset)
        tm_missing   = subset['ТМ'].isna().sum()   + (subset['ТМ']   == '').sum()
        smak_missing = subset['Смак'].isna().sum()  + (subset['Смак'] == '').sum()
        vaga_missing = subset['Вага'].isna().sum()  + (subset['Вага'] == '').sum()
        stats.append({
            'Категорія': cat, 'Підкатегорія': subcat,
            'Всього': total,
            'ТМ не заповнено': tm_missing,
            'Смак не заповнено': smak_missing,
            'Вага не заповнено': vaga_missing,
            'Разом не заповнено': tm_missing + smak_missing + vaga_missing
        })

stats_df = pd.DataFrame(stats)

# Зберігаємо звіт по помилках
LOG_FILE = REF_DIR / f"missing_report_{timestamp}.xlsx"
stats_df.to_excel(LOG_FILE, index=False)
print(f"\nЗвіт збережено: {LOG_FILE.name}")

# Термінал — по кожній колонці
print(f"\nСТАН ДОВІДНИКА")
print(f"{'='*50}")

for col, label in [
    ('ТМ не заповнено', 'ТМ'),
    ('Смак не заповнено', 'Смак'),
    ('Вага не заповнено', 'Вага')
]:
    worst = stats_df.loc[stats_df[col].idxmax()]
    best  = stats_df[stats_df[col] > 0].nsmallest(1, col)

    if worst[col] == 0:
        print(f"\nКолонка [{label}]: 🟢 повністю заповнена")
        continue

    print(f"\nКолонка [{label}]:")
    print(f" 🔴 Найбільше не заповнено: {worst['Категорія']} / {worst['Підкатегорія']} — {worst[col]} шт.")
    if len(best) > 0:
        b = best.iloc[0]
        print(f" 🟡 айменше не заповнено:  {b['Категорія']} / {b['Підкатегорія']} — {b[col]} шт.")

print(f"\n{'='*50}")
worst_total  = stats_df.loc[stats_df['Разом не заповнено'].idxmax()]
filled_total = stats_df[stats_df['Разом не заповнено'] == 0]

print(f"\nЗагально найбільше не заповнено:")
print(f"   {worst_total['Категорія']} / {worst_total['Підкатегорія']} ({worst_total['Всього']} товарів)")
print(f"   ТМ: {worst_total['ТМ не заповнено']}, Смак: {worst_total['Смак не заповнено']}, Вага: {worst_total['Вага не заповнено']}")

if len(filled_total) > 0:
    print(f"\nПовністю заповнені підкатегорії:")
    for _, row in filled_total.iterrows():
        print(f"   {row['Категорія']} / {row['Підкатегорія']}")
else:
    print(f"\nПовністю заповнених підкатегорій поки немає")

# Не визначені товари
master_df['Категорія'] = master_df['Категорія'].str.strip()
master_df['Підкатегорія'] = master_df['Підкатегорія'].str.strip()
cat_missing  = master_df['Категорія'].isna().sum()
subcat_missing = master_df['Підкатегорія'].isna().sum()

print(f"\n{'='*50}")
print(f"Не визначено категорію:    {cat_missing} / {len(master_df)}")
print(f"Не визначено підкатегорію: {subcat_missing} / {len(master_df)}")
print(master_df[master_df['Категорія'].isna() | (master_df['Категорія'] == '')]['товар'].tolist())

print(f"\nMaster файл: {MASTER_FILE}")
print(f"Всього товарів: {len(master_df)}")
print(f"З визначеним ТМ:    {master_df['ТМ'].notna().sum()} / {len(master_df)}")
print(f"З визначеним смаком: {master_df['Смак'].notna().sum()} / {len(master_df)}")
print(f"З визначеною вагою:  {master_df['Вага'].notna().sum()} / {len(master_df)}")