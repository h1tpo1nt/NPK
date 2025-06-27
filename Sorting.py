import pandas as pd
import re
import os
# ======================================
# Настройки путей и параметров
# ======================================
input_file = '/Users/h1tpo1nt/Desktop/test.xlsx'
keywords_file = '/Users/h1tpo1nt/Desktop/keywords.xlsx'
target_col_name = "G31_1 (Описание и характеристика товара)"

# Проверяем существование файлов
if not os.path.exists(input_file):
    raise FileNotFoundError(f"❌ Файл с данными не найден: {input_file}")
if not os.path.exists(keywords_file):
    raise FileNotFoundError(f"❌ Файл с ключами не найден: {keywords_file}")
    
def extract_npk(description, keywords):
    desc = str(description).lower().strip()
    desc = re.sub(r'[\s\xa0\u3000]+', ' ', desc)

    # Формат NPK
    npk_match = re.search(r'\bnpk\s*(?:$s$)?\s*(\d+(?:\.\d+)?)\s*[-:/]\s*(\d+(?:\.\d+)?)\s*[-:/]\s*(\d+(?:\.\d+)?)', desc, re.IGNORECASE)
    if npk_match:
        n = float(npk_match.group(1))
        p = float(npk_match.group(2))
        k = float(npk_match.group(3))
        return {
            'N': int(n) if n.is_integer() else n,
            'P': int(p) if p.is_integer() else p,
            'K': int(k) if k.is_integer() else k
        }

    # Поиск по ключевым словам из Excel
    elements = {'N': 0, 'P': 0, 'K': 0}
    for el in ['N', 'P', 'K']:
        for keyword in keywords[el]:
            pattern = rf'{keyword}\D*?(\d+(?:[,.]\d+)?)%?'
            match = re.search(pattern, desc)
            if match:
                try:
                    value = float(match.group(1).replace(',', '.'))
                    elements[el] = int(value) if value.is_integer() else value
                except ValueError:
                    continue
                break

    return elements



# Чтение ключевых слов
df_keywords = pd.read_excel(keywords_file)
keywords = {
    'N': df_keywords['N'].dropna().astype(str).str.lower().tolist(),
    'P': df_keywords['P'].dropna().astype(str).str.lower().tolist(),
    'K': df_keywords['K'].dropna().astype(str).str.lower().tolist()
}

# Чтение данных
df = pd.read_excel(input_file)

# Нормализуем заголовки для поиска (удаляем лишние пробелы и символы)
normalized_cols = [
    re.sub(r'[\s\xa0\u3000]+', ' ', str(col)).strip() for col in df.columns
]

# Проверяем наличие нужного столбца
if target_col_name not in normalized_cols:
    raise KeyError(f"❌ В таблице отсутствует колонка с точным названием: '{target_col_name}'")

# Находим оригинальное имя столбца
original_target_col = df.columns[normalized_cols.index(target_col_name)]

# Обработка описаний и создание новых колонок
n_values = []
p_values = []
k_values = []
марки = []
npk_flags = []

for idx, row in df.iterrows():
    description = str(row[original_target_col])
    result = extract_npk(description, keywords)

    n = result['N']
    p = result['P']
    k = result['K']

    n_values.append(n)
    p_values.append(p)
    k_values.append(k)

    brand = f"{n}-{p}-{k}"
    марки.append(brand)

    if brand == "0-0-0":
        npk_flags.append("NPK")
    else:
        npk_flags.append("")

# Добавляем новые столбцы в конец DataFrame
df['N'] = n_values
df['P'] = p_values
df['K'] = k_values
df['Марка'] = марки
df['NPK'] = npk_flags

# Записываем результат обратно в исходный лист
with pd.ExcelWriter(input_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name=writer.sheets.keys().__iter__().__next__(), index=False)

print(f"✅ Обработка завершена. Добавлены столбцы 'N', 'P', 'K', 'Марка' и 'NPK' справа от таблицы.")
