import re
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from datetime import datetime

from config import categories, income_categories
from storage import load_data, save_data, update_excel


def detect_category(description):
    description = description.lower()
    
    for category, keywords in categories.items():
        for word in keywords:
            if word in description:# Убираем смайлик для хранения категории в JSON
                return category.split(' ', 1)[1] if ' ' in category else category
    
    return 'Прочее'

def detect_income_category(description):
    """Определение категории дохода"""
    description = description.lower()
    for category, keywords in income_categories.items():
        for word in keywords:
            if word in description:
                return category
    return 'Другое'

def parse_expense(text): #изменённая функция для парсинга с ручным указанием категории в скобках
    match = re.match(r'^(\d+)\s*(.*)', text.strip())
    if not match:
        return None
    amount = int(match.group(1))
    full_description = match.group(2).strip()

    cat_match = re.search(r'\(([^)]+)\)$', full_description)# Проверяем есть ли (категория) в конце
    
    if cat_match:
        manual_cat = cat_match.group(1).lower()  #то что в скобках
        description = full_description[:cat_match.start()].strip()  # текст без скобок
        
        found_category = None# Ищем совпадение с нашими категориями
        for cat in categories:
            cat_name = cat.split(' ', 1)[1].lower()  #убираем смайлик
            if manual_cat in cat_name:
                found_category = cat.split(' ', 1)[1]  #берём без смайлика
                break
        
        if found_category:
            return amount, description, found_category  # возвращаем 3 значения
    
    return amount, full_description, None  # None = определить автоматически

def handle_expense(bot, user_id, text): # изменённая функция для обработки расходов с ручным указанием категории в скобках
    result = parse_expense(text)
    if not result:
        bot.send_message(user_id, "❌ Неверный формат. Введите: 500 кофе")
        return
    
    amount, description, manual_category = result
    
    # Если категория указана вручную используем её, если нет то определяем автоматически
    if manual_category:
        category = manual_category
    else:
        category = detect_category(description)
    
    data = load_data()# Сохранение в json
    if 'expenses' not in data:
        data['expenses'] = {}
    if str(user_id) not in data['expenses']:
        data['expenses'][str(user_id)] = []
    
    data['expenses'][str(user_id)].append({
        'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'amount': amount,
        'description': description,
        'category': category
    })
    save_data(data)
    update_excel(user_id)
    bot.send_message(user_id, f"✅ Записано: {amount}₸ - {description}\n📂 Категория: {category}")

def handle_income(bot, user_id, text):
    match = re.match(r'^\+(\d+)\s*(.*)', text.strip())
    if not match:
        bot.send_message(user_id, "❌ Неверный формат. Введите: +200000 зарплата")
        return

    amount = int(match.group(1))
    description = match.group(2).strip()
    
    category = detect_income_category(description)# Определяем категорию дохода
    data = load_data()# Сохранение
    if 'incomes' not in data:
        data['incomes'] = {}
    if str(user_id) not in data['incomes']:
        data['incomes'][str(user_id)] = []
    
    data['incomes'][str(user_id)].append({
        'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'amount': amount,
        'description': description,
        'category': category
    })
    save_data(data)
    update_excel(user_id) # обновляем Excel отчет после добавления новой записи
    bot.send_message(user_id, f"💰 Записано ДОХОД: {amount}₸ - {description}\n📂 Категория: {category}")

def send_and_remove(bot, user_id, filename): # отправляем фото и удаляем файл
    with open(filename, "rb") as photo:
        bot.send_photo(user_id, photo)
    os.remove(filename)

# Отчёты

def expense_report(bot, user_id):
    """Текстовый отчет за текущий месяц"""
    from storage import get_user_name
    user_name = get_user_name(user_id)

    if not user_name:
        bot.send_message(user_id, "Сначала зарегистрируйтесь через /start")
        return
    
    now = datetime.now()
    current_month = now.strftime('%Y-%m')
    month_name = now.strftime('%B %Y')
    
    data = load_data()
    expenses = data.get('expenses', {}).get(str(user_id), [])
    
    # Фильтруем расходы за текущий месяц
    month_expenses = []
    for exp in expenses:
        if exp['date'].startswith(current_month):
            month_expenses.append(exp)
    
    if not month_expenses:
        bot.send_message(user_id, f"📭 Нет расходов за {month_name}")
        return
    
    # Подсчет по категориям
    total = 0
    categories_sum = {}
    
    for exp in month_expenses:
        amount = exp['amount']
        category = exp['category']
        total += amount
        categories_sum[category] = categories_sum.get(category, 0) + amount
    
    # Сортируем по убыванию
    sorted_cats = sorted(categories_sum.items(), key=lambda x: x[1], reverse=True)
    
    # Формируем отчет
    text = f"📊 *ОТЧЕТ ЗА {month_name.upper()}*\n\n"
    text += f"💰 *Всего расходов:* {total:,}₸\n\n"
    text += "*По категориям:*\n"
    
    for cat, amount in sorted_cats:
        percent = (amount / total) * 100
        text += f"• {cat}: {amount:,}₸ ({percent:.1f}%)\n"
    
    bot.send_message(user_id, text, parse_mode='Markdown')

def income_report(bot, user_id):
    """Отчет по доходам за текущий месяц"""
    from storage import get_user_name
    user_name = get_user_name(user_id)

    if not user_name:
        bot.send_message(user_id, "Сначала зарегистрируйтесь через /start")
        return

    now = datetime.now() #Получаем текущую дату и времяS
    current_month = now.strftime('%Y-%m')
    month_name = now.strftime('%B %Y')

    data = load_data()
    incomes = data.get('incomes', {}).get(str(user_id), [])

    month_incomes = []# фильтр по текущему месяцу
    for inc in incomes:
        if inc['date'].startswith(current_month):
            month_incomes.append(inc)

    if not month_incomes:
        bot.send_message(user_id, f"📭 Нет доходов за {month_name}")
        return

    total = 0
    categories_sum = {}# подсчет по категориям

    for inc in month_incomes: # проходим по каждому доходу
        amount = inc['amount']
        category = inc['category']

        total += amount
        categories_sum[category] = categories_sum.get(category, 0) + amount# сохраняем сумму по каждой категории

    sorted_cats = sorted(categories_sum.items(), key=lambda x: x[1], reverse=True)

    text = f"💰 *ДОХОДЫ ЗА {month_name.upper()}*\n\n"
    text += f"💵 *Всего доходов:* {total:,}₸\n\n"
    text += "*По категориям:*\n"

    for cat, amount in sorted_cats:
        percent = (amount / total) * 100
        text += f"• {cat}: {amount:,}₸ ({percent:.1f}%)\n"

    bot.send_message(user_id, text, parse_mode='Markdown')

def salary_analysis(bot, user_id):
    data = load_data()
    incomes = data.get('incomes', {}).get(str(user_id), [])

    salary_incomes = []

    for inc in incomes:
        if 'Зарплата' in inc['category']:
            salary_incomes.append(inc)

    if not salary_incomes:
        bot.send_message(user_id, "Нет данных по зарплате")
        return

    salary_by_month = {}

    for inc in salary_incomes:
        month = inc['date'][:7]
        amount = inc['amount']

        salary_by_month[month] = salary_by_month.get(month, 0) + amount

    values = list(salary_by_month.values())

    max_month = max(salary_by_month, key=salary_by_month.get)
    max_value = salary_by_month[max_month]

    min_month = min(salary_by_month, key=salary_by_month.get)
    min_value = salary_by_month[min_month]

    avg_salary = sum(values) / len(values)# среднее

    text = "📊 Аналитика зарплат:\n\n"
    text += f"📈 Максимум: {max_value:,}₸ ({max_month})\n"
    text += f"📉 Минимум: {min_value:,}₸ ({min_month})\n"
    text += f"📊 Средняя: {avg_salary:,.0f}₸"

    bot.send_message(user_id, text)

def financial_report(bot, user_id):
    """Финансовый отчет за текущий месяц"""
    from storage import get_user_name
    user_name = get_user_name(user_id)

    if not user_name:
        bot.send_message(user_id, "Сначала зарегистрируйтесь - отправьте /start")
        return

    now = datetime.now()# Получаем текущую дату для определения месяца и года
    current_month = now.strftime('%Y-%m')
    month_name = now.strftime('%B %Y')

    data = load_data() # Загружаем данные и получаем расходы и доходы для пользователя

    expenses = data.get('expenses', {}).get(str(user_id), []) 
    incomes = data.get('incomes', {}).get(str(user_id), [])

    month_expenses = [e for e in expenses if e['date'].startswith(current_month)] #фильтруем по месяцу
    month_incomes = [i for i in incomes if i['date'].startswith(current_month)]

    total_expenses = sum(e['amount'] for e in month_expenses)
    total_incomes = sum(i['amount'] for i in month_incomes)

    balance = total_incomes - total_expenses

    if total_incomes > 0:
        savings_percent = (balance / total_incomes) * 100
    else:
        savings_percent = 0

    text = f"🏦 *ФИНАНСОВЫЙ ОТЧЕТ за {month_name.upper()}*\n\n"

    text += f"💰 *Доходы:* {total_incomes:,}₸\n"
    text += f"💸 *Расходы:* {total_expenses:,}₸\n"
    text += "———————————————\n"

    if balance >= 0:
        text += f"🟢 *Остаток:* +{balance:,}₸\n"
    else:
        text += f"🔴 *Дефицит:* {balance:,}₸\n"

    text += f"\n📊 *Экономия:* {savings_percent:.1f}%"

    bot.send_message(user_id, text, parse_mode='Markdown')

def detail_report(bot, user_id):
    data = load_data()
    expenses = data.get('expenses', {}).get(str(user_id), [])

    if not expenses:
        bot.send_message(user_id, "Нет данных по расходам")
        return

    # Группируем: категория -> описание -> сумма
    report = {}
    for exp in expenses:
        cat = exp['category']
        desc = exp['description'] if exp['description'] else 'без описания'
        amount = exp['amount']

        if cat not in report:
            report[cat] = {}
        report[cat][desc] = report[cat].get(desc, 0) + amount

    text = "📊 *Детальный отчет по категориям:*\n"

    for cat, items in sorted(report.items()):
        cat_total = sum(items.values())
        text += f"\n*{cat}* — итого: {cat_total:,}₸\n"
        # сортируем позиции по убыванию суммы
        for desc, amount in sorted(items.items(), key=lambda x: x[1], reverse=True):
            text += f"  • {desc}: {amount:,}₸\n"

    bot.send_message(user_id, text, parse_mode='Markdown')

# Подписки

def add_subscription(bot, user_id, message_text):
    parts = message_text.strip().split()

    # /add_sub Netflix 2990
    if len(parts) < 3:
        bot.send_message(user_id, "❌ Формат: /add_sub Netflix 2990")
        return

    name = parts[1]       # название подписки
    try:
        price = int(parts[2])
    except ValueError:
        bot.send_message(user_id, "❌ Цена должна быть числом. Пример: /add_sub Netflix 2990")
        return

    data = load_data()

    if 'subscriptions' not in data:
        data['subscriptions'] = {}
    if str(user_id) not in data['subscriptions']:
        data['subscriptions'][str(user_id)] = []

    # Проверяем что такая подписка ещё не добавлена
    existing = [s['name'].lower() for s in data['subscriptions'][str(user_id)]]
    if name.lower() in existing:
        bot.send_message(user_id, f"⚠️ Подписка {name} уже есть в списке")
        return

    data['subscriptions'][str(user_id)].append({
        'name': name,
        'price': price,
        'added_at': datetime.now().strftime('%Y-%m-%d')
    })

    save_data(data)
    bot.send_message(user_id, f"✅ Подписка {name} добавлена — {price:,}₸/месяц")

def remove_subscription(bot, user_id, message_text):
    parts = message_text.strip().split()

    if len(parts) < 2:
        bot.send_message(user_id, "❌ Формат: /remove_sub Netflix")
        return

    name = parts[1].lower()
    data = load_data()
    subs = data.get('subscriptions', {}).get(str(user_id), [])

    new_subs = [s for s in subs if s['name'].lower() != name]

    if len(new_subs) == len(subs):
        bot.send_message(user_id, f"❌ Подписка {parts[1]} не найдена")
        return

    data['subscriptions'][str(user_id)] = new_subs
    save_data(data)
    bot.send_message(user_id, f"🗑 Подписка {parts[1]} удалена")

def show_subscriptions(bot, user_id):
    data = load_data()

    subs = data.get('subscriptions', {}).get(str(user_id), [])

    if not subs:
        bot.send_message(user_id, "📭 Нет подписок. Добавь: /add_sub Netflix 2990")
        return

    # Смотрим расходы за текущий месяц
    current_month = datetime.now().strftime('%Y-%m')
    expenses = data.get('expenses', {}).get(str(user_id), [])

    # Собираем названия трат за этот месяц в нижнем регистре
    paid_this_month = set()
    for exp in expenses:
        if exp['date'].startswith(current_month):
            paid_this_month.add(exp['description'].lower())

    paid = []
    unpaid = []
    total_price = sum(s['price'] for s in subs)

    for sub in subs:
        if sub['name'].lower() in paid_this_month:
            paid.append(sub)
        else:
            unpaid.append(sub)

    text = f"📺 *Подписки за {datetime.now().strftime('%B %Y')}*\n\n"

    if unpaid:
        text += "❌ *Не оплачены:*\n"
        for s in unpaid:
            text += f"  • {s['name']}: {s['price']:,}₸\n"

    if paid:
        text += "\n✅ *Оплачены:*\n"
        for s in paid:
            text += f"  • {s['name']}: {s['price']:,}₸\n"

    text += f"\n———————————————\n"
    text += f"💸 *Итого подписок:* {total_price:,}₸/мес"

    bot.send_message(user_id, text, parse_mode='Markdown')

# Графики

def balance_trend(bot, user_id):
    data = load_data()
    expenses = data.get('expenses', {}).get(str(user_id), [])
    incomes = data.get('incomes', {}).get(str(user_id), [])

    if not expenses and not incomes:
        bot.send_message(user_id, "Нет данных")
        return

    # расходы
    exp_by_month = {}
    for exp in expenses:
        month = exp['date'][:7]
        exp_by_month[month] = exp_by_month.get(month, 0) + exp['amount']

    # доходы
    inc_by_month = {}
    for inc in incomes:
        month = inc['date'][:7]
        inc_by_month[month] = inc_by_month.get(month, 0) + inc['amount']

    # все месяцы
    all_months = sorted(set(exp_by_month) | set(inc_by_month))

    balances = []

    for m in all_months:
        income = inc_by_month.get(m, 0)
        expense = exp_by_month.get(m, 0)
        balance = income - expense
        balances.append(balance)

    plt.figure()
    plt.plot(all_months, balances, marker='o')

    plt.axhline(0)  # линия нуля

    plt.title("Финансовый остаток по месяцам")
    plt.xlabel("Месяц")
    plt.ylabel("Остаток (₸)")

    # подписи
    for i, v in enumerate(balances):
        plt.text(i, v, str(v))

    plt.savefig("balance.png")
    plt.close()

    send_and_remove(bot, user_id, "balance.png")

    total_saved = 0 # считаем общий остаток за все месяцы 
    text = "💰 *Остаток по месяцам:*\n\n"

    for i, m in enumerate(all_months):
        income = inc_by_month.get(m, 0)
        expense = exp_by_month.get(m, 0)
        saved = income - expense
        total_saved += saved
        percent = (saved / income * 100) if income > 0 else 0
        emoji = "🟢" if saved >= 0 else "🔴"
        text += f"{emoji} {m}: {saved:,}₸ ({percent:.1f}%)\n"

    text += f"\n———————————————\n"
    text += f"💵 *Итого накоплено:* {total_saved:,}₸"

    bot.send_message(user_id, text, parse_mode='Markdown')

def all_categories_chart(bot, user_id):
    data = load_data()
    expenses = data.get('expenses', {}).get(str(user_id), [])

    if not expenses:
        bot.send_message(user_id, "Нет данных по расходам")
        return

    # группировка по категориям
    categories_sum = {}
    total = 0

    for exp in expenses:
        cat = exp['category']
        amount = exp['amount']

        categories_sum[cat] = categories_sum.get(cat, 0) + amount
        total += amount

    labels = list(categories_sum.keys())
    values = list(categories_sum.values())
    
    plt.figure(figsize=(10, 8))
    
    def autopct(pct):# функция для отображения процентов + суммы
        absolute = int(pct * total / 100)
        return f"{pct:.1f}%\n({absolute}₸)"
    
    plt.pie(
        values,
        labels=labels,
        autopct=autopct,
        pctdistance=0.75,# цифры % ближе к центру
        labeldistance=1.2,# подписи дальше от края
    )

    plt.title("Распределение расходов за всё время")

    plt.savefig("pie.png")
    plt.close()

    send_and_remove(bot, user_id, "pie.png")

def compare_chart(bot, user_id):
    data = load_data()
    expenses = data.get('expenses', {}).get(str(user_id), [])
    incomes = data.get('incomes', {}).get(str(user_id), [])

    if not expenses and not incomes:
        bot.send_message(user_id, "Нет данных")
        return

    # расходы
    exp_by_month = {}
    for exp in expenses:
        month = exp['date'][:7]
        exp_by_month[month] = exp_by_month.get(month, 0) + exp['amount']

    # доходы
    inc_by_month = {}
    for inc in incomes:
        month = inc['date'][:7]
        inc_by_month[month] = inc_by_month.get(month, 0) + inc['amount']

    # все месяцы
    all_months = sorted(set(exp_by_month) | set(inc_by_month))# set собираем уникальные месяцы

    exp_values = [exp_by_month.get(m, 0) for m in all_months]
    inc_values = [inc_by_month.get(m, 0) for m in all_months]

    x = range(len(all_months))
    width = 0.4 # 

    plt.figure()

    plt.bar([i - width/2 for i in x], inc_values, width=width, label="Доходы")
    plt.bar([i + width/2 for i in x], exp_values, width=width, label="Расходы")

    plt.xticks(x, all_months, rotation=45)
    plt.title("Доходы vs Расходы")
    plt.xlabel("Месяц")
    plt.ylabel("Сумма (₸)")
    plt.legend()

    plt.tight_layout()
    plt.savefig("compare.png")
    plt.close()

    send_and_remove(bot, user_id, "compare.png")

def monthly_chart(bot, user_id):
    current_month = datetime.now().strftime('%Y-%m')

    data = load_data()
    expenses = data.get('expenses', {}).get(str(user_id), [])

    month_expenses = [exp for exp in expenses if exp['date'].startswith(current_month)] # оставляем только расходы за текущий месяц

    if not month_expenses:
        bot.send_message(user_id, "Нет расходов за этот месяц")
        return

    categories_sum = {}# группировка
    for exp in month_expenses:
        cat = exp['category']
        amount = exp['amount']
        categories_sum[cat] = categories_sum.get(cat, 0) + amount

    labels = list(categories_sum.keys())# данные для графика
    values = list(categories_sum.values())

    plt.figure()
    plt.barh(labels, values)

    plt.title("Расходы за текущий месяц")
    plt.xlabel("Сумма (₸)")
    plt.ylabel("Категории")

    plt.savefig("chart.png")
    plt.close()

    send_and_remove(bot, user_id, "chart.png")