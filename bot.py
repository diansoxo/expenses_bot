import telebot
from config import TOKEN, categories, income_categories
from storage import get_user_name, register_user
from services import (
    handle_expense, handle_income, send_and_remove,
    expense_report, income_report, salary_analysis,
    financial_report, detail_report,
    add_subscription, remove_subscription, show_subscriptions,
    balance_trend, all_categories_chart, compare_chart, monthly_chart
)

bot = telebot.TeleBot(TOKEN)

waiting_for_name = set()


@bot.message_handler(commands=['start'])
def start(message):
    user_id = message.from_user.id
    user_name = get_user_name(user_id)
    if not user_name:
        waiting_for_name.add(user_id)
        bot.send_message(user_id, "Добро пожаловать! Как я могу к Вам обращаться?")
    else:
        bot.send_message(user_id, f"С возвращением, {user_name}!...")
    text = (
        "👋 Привет!Я бот для учета расходов и доходов\n\n"

        "📌 Как пользоваться:\n"
        "Просто отправь сообщение в формате:\n"
        "👉 1500 кофе\n\n"

        "💰 Для записи дохода:\n"
        "👉 +200000 зарплата\n\n"

        "📂 Я автоматически определю категорию:\n"
        "🍔 Еда- обед, кофе, пицца\n"
        "🚇 Транспорт- такси, метро, автобус\n"
        "🎮 Развлечения- кино, игры, концерт\n"
        "💊 Здоровье- аптека, врач\n"
        "🏠 Жилье- аренда\n"
        "📱 Связь- интернет, телефон\n"
        "💄 Красота- маникюр, стрижка\n"
        "🛒 Покупки- одежда, джинсы\n"
        "💸 Прочее- всё остальное\n\n"

        "📊 Доступные команды:\n"
        "/categories- категории\n"#есть
        "/monthly- график расходов\n"
        "/income_report- доходы\n"
        "/expense_report- отчет за месяц\n"#есть
        "/financial_report- доходы vs расходы\n"
        "/compare- доходы vs расходы\n"
        "/salary_analysis - аналитика зарплаты\n"
        "/balance_trend- остаток по месяцам\n"
        "/detail_report- детальный отчет\n"
        "/subscriptions- статус подписок за месяц\n"
        "/all_categories- круговая диаграмма расходов за все время\n\n"

        "🚀 Просто отправь сумму и описание!"
    )
    bot.send_message(message.chat.id, text)

@bot.message_handler(commands=['change_name'])
def change_name(message):
    user_id = message.from_user.id
    waiting_for_name.add(user_id)
    bot.send_message(user_id, "Введите новое имя:")

@bot.message_handler(commands=['categories'])
def show_categories(message):
    """Показать все категории"""
    text = "📂 Категории расходов:\n"
    for cat, keywords in categories.items():
        text += f"\n• {cat}: {', '.join(keywords[:3])}"
        if len(keywords) > 3:
            text += f" +{len(keywords)-3}"
    
    text += "\n\n💰 Категории доходов:\n"
    for cat, keywords in income_categories.items():
        if keywords:
            text += f"\n• {cat}: {', '.join(keywords)}"
    
    bot.send_message(message.chat.id, text)

@bot.message_handler(commands=['expense_report'])
def cmd_expense_report(message):
    expense_report(bot, message.from_user.id)

@bot.message_handler(commands=['income_report'])
def cmd_income_report(message):
    income_report(bot, message.from_user.id)

@bot.message_handler(commands=['salary_analysis'])
def cmd_salary_analysis(message):
    salary_analysis(bot, message.from_user.id)

@bot.message_handler(commands=['financial_report'])
def cmd_financial_report(message):
    financial_report(bot, message.from_user.id)

@bot.message_handler(commands=['detail_report'])
def cmd_detail_report(message):
    detail_report(bot, message.from_user.id)

@bot.message_handler(commands=['add_sub'])
def cmd_add_sub(message):
    add_subscription(bot, message.from_user.id, message.text)

@bot.message_handler(commands=['remove_sub'])
def cmd_remove_sub(message):
    remove_subscription(bot, message.from_user.id, message.text)

@bot.message_handler(commands=['subscriptions'])
def cmd_subscriptions(message):
    show_subscriptions(bot, message.from_user.id)

# Гафики

@bot.message_handler(commands=['balance_trend'])
def cmd_balance_trend(message):
    balance_trend(bot, message.from_user.id)

@bot.message_handler(commands=['all_categories'])
def cmd_all_categories(message):
    all_categories_chart(bot, message.from_user.id)

@bot.message_handler(commands=['compare'])
def cmd_compare(message):
    compare_chart(bot, message.from_user.id)

@bot.message_handler(commands=['monthly'])
def cmd_monthly(message):
    monthly_chart(bot, message.from_user.id)

@bot.message_handler(func=lambda message: True)
def handle_message(message):
    user_id = message.from_user.id
    text = message.text.strip()
    
    if user_id in waiting_for_name:
        waiting_for_name.discard(user_id)#
        register_user(user_id, text)
        bot.send_message(user_id, f"Отлично, {text}! Теперь можете пользоваться ботом.")
        return
    
    user_name = get_user_name(user_id)
    if not user_name:
        waiting_for_name.add(user_id)
        bot.send_message(user_id, "Добро пожаловать! Как я могу к Вам обращаться?")
        return

    if text.startswith('+'):
        handle_income(bot, user_id, text)
    else:
        handle_expense(bot, user_id, text)

bot.infinity_polling()