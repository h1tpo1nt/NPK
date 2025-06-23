import pandas as pd
import re
from openpyxl import load_workbook

def extract_npk(description):
    desc = str(description).lower().strip()

 # Сначала проверяем формат NPK, например: "NPK 19-4-19" или "NPK 13:19:19"
 npk_match = re.search(r'npk\s*(\d+)[\-\:\s]*(\d+)[\-\:\s]*(\d+)', desc, re.IGNORECASE)
    if npk_match:
        return {
            'N': int(npk_match.group(1)),
            'P': int(npk_match.group(2)),
            'K': int(npk_match.group(3))
        }

 # Словарь элементов с ключевыми словами и позициями
 elements = {
        'N': {'keyword': r'азот', 'pos': float('inf'), 'value': 0},
        'P': {'keyword': r'фосфор', 'pos': float('inf'), 'value': 0},
        'K': {'keyword': r'кали[й]', 'pos': float('inf'), 'value': 0}
    }

  # Поиск позиций каждого слова в строке
 for el_key, data in elements.items():
        match = re.search(data['keyword'], desc, re.IGNORECASE)
        if match:
            elements[el_key]['pos'] = match.start()

 # Сортировка по позиции
  sorted_elements = sorted(elements.items(), key=lambda x: x[1]['pos'])
    present_elements = [item for item in sorted_elements if item[1]['pos'] != float('inf')]

   # Извлечение чисел между элементами
   for i, (key, data) in enumerate(present_elements):
        current_pos = data['pos']
        next_pos = len(desc) if i == len(present_elements) - 1 else present_elements[i + 1][1]['pos']

   segment = desc[current_pos:next_pos]
        num_match = re.search(r'\d+[\.,]?\d*', segment)

 if num_match:
            try:
                value = float(num_match.group().replace(',', '.'))
                elements[key]['value'] = value
            except ValueError:
                elements[key]['value'] = 0
        else:
            elements[key]['value'] = 0

return {
        'N': elements['N']['value'],
        'P': elements['P']['value'],
        'K': elements['K']['value']
    }

# Путь к файлу
input_file = 'input.xlsx'

# Чтение данных из текущего листа
df = pd.read_excel(input_file)

# Добавляем новые столбцы
df['N'] = 0
df['P'] = 0
df['K'] = 0

# Обработка каждой строки
for idx, row in df.iterrows():
    result = extract_npk(row['G31_1'])
    df.at[idx, 'N'] = result['N']
    df.at[idx, 'P'] = result['P']
    df.at[idx, 'K'] = result['K']

# Записываем результат в тот же файл, но на новый лист
with pd.ExcelWriter(input_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Output', index=False)

print(f"✅ Обработка завершена. Результат записан на лист 'Output' в файл: {input_file}")
