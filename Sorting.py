import pandas as pd
import re

def extract_npk(description):
    desc = str(description).lower().strip()

    # Сначала проверяем формат NPK
    npk_match = re.search(r'npk\s*(\d+)[\-\:\s]*(\d+)[\-\:\s]*(\d+)', desc, re.IGNORECASE)
    if npk_match:
        return {
            'N': int(npk_match.group(1)),
            'P': int(npk_match.group(2)),
            'K': int(npk_match.group(3))
        }

    # Словарь элементов с гибкими паттернами
    elements = {
        'N': {'root': r'\bазот', 'value': 0},
        'P': {'root': r'\bфосфор|\bp2o5|\bп2о5', 'value': 0},
        'K': {'root': r'\bкали[йяие]|k2o', 'value': 0}
    }

    # Находим все совпадения ключевых слов
    all_matches = []
    for key, data in elements.items():
        matches = re.finditer(data['root'], desc, re.IGNORECASE)
        for match in matches:
            all_matches.append((match.start(), match.end(), key))

    # Сортируем по позиции
    all_matches.sort()

    # Извлечение чисел между элементами
    for i, (start, end, key) in enumerate(all_matches):
        next_start = len(desc) if i == len(all_matches) - 1 else all_matches[i + 1][0]
        segment = desc[end:next_start]

        num_match = re.search(r'(\d+[\.,]?\d*)\s*(?:\+/\-[\s\d\.\,%]*)?', segment)

        if num_match:
            try:
                value = float(num_match.group(1).replace(',', '.'))
                elements[key]['value'] = value
            except ValueError:
                pass

    return {
        'N': elements['N']['value'],
        'P': elements['P']['value'],
        'K': elements['K']['value']
    }

# Путь к файлу
input_file = '/Users/h1tpo1nt/Desktop/test.xlsx'

# Чтение данных из исходного листа
df = pd.read_excel(input_file)

# Починить заголовки
df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]

# Проверка наличия нужной колонки
if 'G31_1' not in df.columns:
    raise KeyError("В таблице отсутствует колонка 'G31_1'")

# Обработка описаний и создание колонки "Марка"
марки = []
npk_flags = []

for idx, row in df.iterrows():
    description = str(row['G31_1'])
    result = extract_npk(description)
    
    n = int(result['N']) if result['N'] != 0 else 0
    p = int(result['P']) if result['P'] != 0 else 0
    k = int(result['K']) if result['K'] != 0 else 0
    
    brand = f"{n}-{p}-{k}"
    марки.append(brand)
    
    if brand == "0-0-0":
        npk_flags.append("NPK")
    else:
        npk_flags.append("")

# Добавляем новые колонки
df['Марка'] = марки
df['NPK'] = npk_flags

# Удаляем временные столбцы N, P, K, если они вдруг появились
cols_to_keep = ['G31_1', 'Марка', 'NPK']
df = df[cols_to_keep]

# Перезаписываем исходный лист
with pd.ExcelWriter(input_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name=writer.sheets.keys().__iter__().__next__(), index=False)

print(f"✅ Обработка завершена. Оставлены только колонки: G31_1, Марка, NPK.")
