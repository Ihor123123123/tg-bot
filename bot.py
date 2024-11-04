import telebot
import speech_recognition as sr
import pandas as pd
from datetime import datetime
import os
from pydub import AudioSegment
import re

# Токен бота
TOKEN = '7917493776:AAHXc2aoYTdKKldmottOkVKZcJvqij7gQkw'
bot = telebot.TeleBot(TOKEN)

# Файл для хранения данных
file_name = "expenses.xlsx"

# Проверяем, существует ли файл; если нет, создаем его
if not os.path.exists(file_name):
    df = pd.DataFrame(columns=["Сумма", "Тип траты", "Описание", "Дата"])
    df.to_excel(file_name, index=False)

# Обработка голосовых сообщений
@bot.message_handler(content_types=['voice'])
def handle_voice(message):
    file_info = bot.get_file(message.voice.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    # Сохраняем голосовое сообщение как файл .ogg
    with open("voice.ogg", 'wb') as new_file:
        new_file.write(downloaded_file)

    # Конвертируем файл в .wav для распознавания
    sound = AudioSegment.from_ogg("voice.ogg")
    sound.export("voice.wav", format="wav")

    # Преобразуем голос в текст
    recognizer = sr.Recognizer()
    with sr.AudioFile("voice.wav") as source:
        audio = recognizer.record(source)
    try:
        text = recognizer.recognize_google(audio, language="ru-RU")
        bot.reply_to(message, f"Распознанный текст: {text}")
        parse_and_save_expense(text, message)
    except sr.UnknownValueError:
        bot.reply_to(message, "Извините, не могу распознать голос.")
    except sr.RequestError:
        bot.reply_to(message, "Сервис распознавания временно недоступен.")

# Функция для анализа текста и записи в Excel
def parse_and_save_expense(text, message):
    try:
        # Определение типа траты: заработок или трата
        if any(keyword in text.lower() for keyword in ["заработал", "доход", "прибыль"]):
            expense_type = "заработок"
        else:
            expense_type = "трата"

        # Извлечение суммы с копейками
        amount_pattern = re.search(r"(\d+)\s*злотых(?:\s*и\s*(\d+)\s*копеек)?", text)
        if amount_pattern:
            zloty = int(amount_pattern.group(1))
            kopeck = int(amount_pattern.group(2)) if amount_pattern.group(2) else 0
            amount = zloty + kopeck / 100
        else:
            bot.send_message(message.chat.id, "Не удалось распознать сумму.")
            return

        # Описание траты (убираем слова, связанные с суммой и типом)
        description = re.sub(r"(\d+\s*злотых(?:\s*и\s*\d+\s*копеек)?|заработал|доход|прибыль|трата|потратил)", "", text, flags=re.IGNORECASE).strip()

        # Дата в формате год-месяц-число
        date = datetime.now().strftime("%Y-%m-%d")

        # Записываем данные в Excel
        new_data = {"Сумма": amount, "Тип траты": expense_type, "Описание": description, "Дата": date}
        df = pd.read_excel(file_name)
        df = df.append(new_data, ignore_index=True)
        df.to_excel(file_name, index=False)

        # Отправляем сообщение с подтверждением
        bot.send_message(
            message.chat.id,
            f"Готово ✅\nСумма: {amount}\nТип траты: {expense_type}\nОписание: {description}\nДата: {date}"
        )

    except Exception as e:
        print(f"Ошибка: {e}")
        bot.send_message(message.chat.id, "Произошла ошибка при обработке сообщения.")

# Запуск бота
bot.polling()
