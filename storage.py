import json
import os
import pandas as pd
from datetime import datetime
from config import DATA_FILE, BASE_DIR


def load_data():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_user_name(user_id):
    """Получить имя пользователя из базы"""
    data = load_data()
    if 'users' not in data:
        data['users'] = {}
    
    if str(user_id) in data['users']:
        return data['users'][str(user_id)].get('name')
    return None

def register_user(user_id, name):
    data = load_data()

    if 'users' not in data:
        data['users'] = {}

    data['users'][str(user_id)] = {
        'name': name,
        'registered_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    save_data(data)

def get_available_months(user_id):
    data = load_data()
    expenses = data.get('expenses', {}).get(str(user_id), [])
    incomes = data.get('incomes', {}).get(str(user_id), [])
    months = set()
    for exp in expenses:
        months.add(exp['date'][:7])
    for inc in incomes:
        months.add(inc['date'][:7])
    return sorted(months, reverse=True)

def update_excel(user_id): # обновление Excel отчета для пользователя
    data = load_data()
    expenses = data.get('expenses', {}).get(str(user_id), [])
    incomes = data.get('incomes', {}).get(str(user_id), [])

    filename = os.path.join(BASE_DIR, f"report_{user_id}.xlsx")

    rows_exp = []
    for exp in expenses:
        dt = datetime.strptime(exp['date'], '%Y-%m-%d %H:%M:%S')
        rows_exp.append({
            'Дата':         dt.strftime('%Y-%m-%d'),
            'Время':        dt.strftime('%H:%M:%S'),
            'Категория':    exp['category'],
            'Наименование': exp['description'],
            'Сумма (₸)':    exp['amount']
        })

    rows_inc = []
    for inc in incomes:
        dt = datetime.strptime(inc['date'], '%Y-%m-%d %H:%M:%S')
        rows_inc.append({
            'Дата':         dt.strftime('%Y-%m-%d'),
            'Время':        dt.strftime('%H:%M:%S'),
            'Категория':    inc['category'],
            'Наименование': inc['description'],
            'Сумма (₸)':    inc['amount']
        })

    df_exp = pd.DataFrame(rows_exp) if rows_exp else pd.DataFrame()
    df_inc = pd.DataFrame(rows_inc) if rows_inc else pd.DataFrame()

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Общий лист расходов
        if not df_exp.empty:
            df_exp.sort_values(['Категория', 'Дата']).to_excel(writer, sheet_name='Все расходы', index=False)
            # Отдельный лист на каждую категорию расходов
            for cat, group in df_exp.groupby('Категория'):
                group.to_excel(writer, sheet_name=cat[:31], index=False)

        # Лист доходов
        if not df_inc.empty:
            df_inc.sort_values(['Дата']).to_excel(writer, sheet_name='Доходы', index=False)