"""
Добавляет колонку "Ответственный" в deals_for_review.csv
Берёт данные из Битрикс24 через вебхук
"""
import csv, os, requests, re, time
from pathlib import Path

WEBHOOK = os.environ['BITRIX_WEBHOOK']  # https://<портал>/rest/<id>/<токен>/
INPUT  = 'deals_for_review.csv'
OUTPUT = 'deals_for_review.csv'  # перезаписываем

# ── Шаг 1: читаем CSV, собираем ID сделок ────────────────────────────────────
rows = []
with open(INPUT, encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    fieldnames = reader.fieldnames
    for row in reader:
        rows.append(row)

deal_ids = []
for row in rows:
    m = re.search(r'/deal/details/(\d+)/', row.get('Ссылка', ''))
    if m:
        deal_ids.append(int(m.group(1)))

print(f'Всего сделок: {len(deal_ids)}')

# ── Шаг 2: получаем ASSIGNED_BY_ID пакетами по 50 ────────────────────────────
id_to_assigned = {}

for i in range(0, len(deal_ids), 50):
    batch_ids = deal_ids[i:i+50]
    params = {
        'filter': {'ID': batch_ids},
        'select': ['ID', 'ASSIGNED_BY_ID'],
        'start': 0
    }
    r = requests.post(WEBHOOK + 'crm.deal.list.json', json=params, timeout=30)
    r.raise_for_status()
    data = r.json()
    for deal in data.get('result', []):
        id_to_assigned[int(deal['ID'])] = int(deal['ASSIGNED_BY_ID'])
    print(f'  Загружено сделок: {i + len(batch_ids)}/{len(deal_ids)}')
    time.sleep(0.3)

# ── Шаг 3: получаем имена пользователей ──────────────────────────────────────
user_ids = list(set(id_to_assigned.values()))
user_names = {}

for uid in user_ids:
    r = requests.get(WEBHOOK + 'user.get.json', params={'ID': uid}, timeout=30)
    if r.status_code != 200:
        print(f'  Пропущен user {uid}: {r.status_code}')
        continue
    result = r.json().get('result', [])
    if result:
        user = result[0]
        user_names[uid] = f"{user.get('NAME', '')} {user.get('LAST_NAME', '')}".strip()
    time.sleep(0.1)

print(f'Загружено пользователей: {len(user_names)}')

# ── Шаг 4: добавляем колонку и сохраняем ─────────────────────────────────────
new_fieldnames = ['Ответственный'] + list(fieldnames)

with open(OUTPUT, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=new_fieldnames)
    writer.writeheader()
    for row in rows:
        m = re.search(r'/deal/details/(\d+)/', row.get('Ссылка', ''))
        deal_id = int(m.group(1)) if m else None
        assigned_id = id_to_assigned.get(deal_id)
        row['Ответственный'] = user_names.get(assigned_id, '') if assigned_id else ''
        writer.writerow(row)

print(f'Готово! Сохранено: {OUTPUT}')

# ── Статистика ────────────────────────────────────────────────────────────────
from collections import Counter
counts = Counter(
    user_names.get(id_to_assigned.get(
        int(re.search(r'/(\d+)/', r['Ссылка']).group(1))
    ), 'Неизвестно')
    for r in rows if re.search(r'/(\d+)/', r.get('Ссылка', ''))
)
print('\nСделок по ответственным:')
for name, cnt in sorted(counts.items(), key=lambda x: -x[1]):
    print(f'  {name}: {cnt}')
