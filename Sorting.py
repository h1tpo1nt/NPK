import pandas as pd
import re

def extract_npk(description):
    desc = str(description).lower().strip()
    desc = re.sub(r'[\s\xa0\u3000]+', ' ', desc)

    # Формат NPK
    npk_match = re.search(r'\bnpk\s*(?:\(s\))?\s*(\d+(?:\.\d+)?)\s*[-:/]\s*(\d+(?:\.\d+)?)\s*[-:/]\s*(\d+(?:\.\d+)?)', desc, re.IGNORECASE)
    if npk_match:
        n = float(npk_match.group(1))
        p = float(npk_match.group(2))
        k = float(npk_match.group(3))
        # Проверяем, являются ли значения целыми
        return {
            'N': int(n) if n.is_integer() else n,
            'P': int(p) if p.is_integer() else p,
            'K': int(k) if k.is_integer() else k
        }

    # Словарь элементов
    elements = {
        'N': {'keywords': [r'\bазот'], 'value': 0},
        'P': {'keywords': [r'\bфосфор', r'\bp2o5', r'\bп2о5'], 'value': 0},
        'K': {'keywords': [r'\bкали[йяие]', r'\bk2o'], 'value': 0}
    }

    for el_key, data in elements.items():
        for keyword in data['keywords']:
            pattern = rf'{keyword}\D*?(\d+(?:[,.]\d+)?)%?'
            match = re.search(pattern, desc)
            if match:
                try:
                    value = float(match.group(1).replace(',', '.'))
                    data['value'] = int(value) if value.is_integer() else value
                except ValueError:
                    continue
                break

    return {
        'N': elements['N']['value'],
        'P': elements['P']['value'],
        'K': elements['K']['value']
    }

def data_to_return(value):
    """Возвращает целое число, если возможно"""
    try:
        return int(value) if value == int(value) else value
    except:
        return 0


# Путь к файлу
input_file = '/Users/h1tpo1nt/Desktop/test.xlsx'

# Точное имя нужного столбца
target_col_name = "G31_1 (Описание и характеристика товара)"

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
марки = []
npk_flags = []

for idx, row in df.iterrows():
    description = str(row[original_target_col])
    result = extract_npk(description)

    n = result['N']
    p = result['P']
    k = result['K']

    brand = f"{n}-{p}-{k}"
    марки.append(brand)

    if brand == "0-0-0":
        npk_flags.append("NPK")
    else:
        npk_flags.append("")

# Добавляем новые столбцы в конец DataFrame
df['Марка'] = марки
df['NPK'] = npk_flags

# Записываем результат обратно в исходный лист
with pd.ExcelWriter(input_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name=writer.sheets.keys().__iter__().__next__(), index=False)

print(f"✅ Обработка завершена. Добавлены столбцы 'Марка' и 'NPK' справа от таблицы.")
