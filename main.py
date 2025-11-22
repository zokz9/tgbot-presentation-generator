import telebot
from telebot.types import Message, InlineKeyboardMarkup, InlineKeyboardButton
from ppt_generator import PresentationGenerator
from config import TELEGRAM_BOT_TOKEN
import json

bot = telebot.TeleBot(TELEGRAM_BOT_TOKEN)
gen = PresentationGenerator(templates_dir="./presentations")

# Хранилище временных данных
user_data = {}

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message: Message):
    welcome_text = """
🤖 *Привет! Я бот для генерации презентаций*

Я могу создать для вас презентацию PowerPoint на любую тему с помощью ИИ Kimi K2.

*Команды:*
/create - Создать презентацию (с шаблоном или с нуля)
/cancel - Отменить действие

Просто напишите /create!
    """
    bot.reply_to(message, welcome_text, parse_mode="Markdown")

@bot.message_handler(commands=['create'])
def start_creation(message: Message):
    """Начинает создание презентации с выбором шаблона"""
    templates = gen.get_available_templates()
    
    text = "📂 *Выберите шаблон* или создайте презентацию без шаблона:\n\n"
    markup = InlineKeyboardMarkup()
    
    # Добавляем кнопки шаблонов
    for idx, tmpl in enumerate(templates, 1):
        text += f"{idx}. {tmpl['name']}\n"
        markup.add(InlineKeyboardButton(
            f"{idx}. {tmpl['name']}", 
            callback_data=f"template_{tmpl['name']}"
        ))
    
    # Добавляем кнопку "Без шаблона"
    markup.add(InlineKeyboardButton(
        "🎨 Без шаблона", 
        callback_data="template_none"
    ))
    
    bot.reply_to(message, text, parse_mode="Markdown", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith("template_"))
def select_template_or_scratch(call):
    """Выбирает шаблон или создает без шаблона"""
    user_id = call.from_user.id
    template_name = call.data.split("_", 1)[1]
    
    if template_name == "none":
        # Создание без шаблона
        user_data[user_id] = {"step": "topic"}
        bot.edit_message_text(
            "📝 Введите тему презентации:",
            call.message.chat.id,
            call.message.message_id
        )
    else:
        # Создание с шаблоном
        user_data[user_id] = {
            "template": template_name,
            "step": "topic"
        }
        
        bot.edit_message_text(
            f"✅ Выбран шаблон: *{template_name}*\n\nТеперь введите тему презентации:",
            call.message.chat.id,
            call.message.message_id,
            parse_mode="Markdown"
        )

@bot.message_handler(func=lambda m: user_data.get(m.from_user.id, {}).get("step") == "topic")
def get_topic(message: Message):
    user_id = message.from_user.id
    user_data[user_id]["topic"] = message.text
    
    template = user_data[user_id].get("template")
    if template:
        bot.reply_to(message, f"✅ Тема: {message.text}")
    
    markup = InlineKeyboardMarkup()
    markup.row(
        InlineKeyboardButton("3 слайда", callback_data="slides_3"),
        InlineKeyboardButton("5 слайдов", callback_data="slides_5"),
        InlineKeyboardButton("7 слайдов", callback_data="slides_7")
    )
    
    bot.send_message(message.chat.id, "📊 Сколько слайдов нужно?", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith("slides_"))
def get_slides_count(call):
    user_id = call.from_user.id
    count = int(call.data.split("_")[1])
    user_data[user_id]["slides_count"] = count
    
    # Тут же генерируем
    generate_presentation(call)

def generate_presentation(call):
    user_id = call.from_user.id
    
    topic = user_data[user_id]["topic"]
    slides_count = user_data[user_id]["slides_count"]
    template_name = user_data[user_id].get("template")
    
    # Путь к шаблону
    template_path = None
    template_structure = None
    
    if template_name:
        template_path = f"./presentations/{template_name}.pptx"
        if template_path:
            template_structure, _ = gen.extract_template_info(template_path)
            if template_structure:
                slides_count = len(template_structure["slides"])
                user_data[user_id]["slides_count"] = slides_count
    
    bot.edit_message_text(
        f"⏳ *Генерация презентации...*\nТема: {topic}\nСлайдов: {slides_count}\nШаблон: {template_name or 'нет'}",
        call.message.chat.id,
        call.message.message_id,
        parse_mode="Markdown"
    )
    
    try:
        # Генерируем структуру
        structure = gen.generate_structure(topic, "ru", slides_count, template_structure)
        
        # Создаем презентацию (в памяти!)
        pptx_buffer = gen.create_presentation(structure, "modern", template_path)
        
        # Отправляем файл
        bot.send_document(
            call.message.chat.id,
            pptx_buffer,
            visible_file_name=f"{topic[:30]}_{template_name or 'new'}.pptx",
            caption="✅ Презентация готова!"
        )
        
    except Exception as e:
        bot.send_message(call.message.chat.id, f"❌ Ошибка: {str(e)}")
    
    finally:
        user_data.pop(user_id, None)

@bot.message_handler(commands=['cancel'])
def cancel_creation(message: Message):
    user_id = message.from_user.id
    user_data.pop(user_id, None)
    bot.reply_to(message, "✅ Отменено.")

if __name__ == "__main__":
    print("🤖 Бот запущен!")
    bot.infinity_polling()
