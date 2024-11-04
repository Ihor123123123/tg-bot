import telebot
import spacy
import pandas as pd
from datetime import datetime
import calendar
import os
from pydub import AudioSegment
import speech_recognition as sr
import re

# Токен бота
TOKEN = '7917493776:AAHXc2aoYTdKKldmottOkVKZcJvqij7gQkw'
bot = telebot.TeleBot(TOKEN)

# Загрузка модели spaCy для русского языка
nlp = spacy.load("ru_core_news_sm")

# Файл для хранения данных
file_name = "expenses.xlsx"

# Список слов-паразитов
stop_words = ["ну", "вот", "я", "а я", "а", "ну вот", "ну я"]

# История записей для отслеживания последних операций
history = []

# Функция для обновления Excel-файла с учетом месячной структуры
def update_monthly_expenses(year, month, day, amount, expense_type, description):
    num_days = calendar.monthrange(year, month)[1]
    print(f"Обновление данных в листе: {calendar.month_name[month]} {year}, количество дней: {num_days}")

    if os.path.exists(file_name):
        df = pd.read_excel(file_name, sheet_name=None)
        print("Файл Excel успешно загружен.")
    else:
        df = {}
        print("Файл Excel не найден. Создаю новый файл.")

    month_name = f"{calendar.month_name[month]}"
    
    if month_name in df:
        month_df = df[month_name]
        print(f"Лист {month_name} найден в файле.")
    else:
        print(f"Лист {month_name} не найден. Создаю новый лист.")
        days = [f"{day:02}.{month:02}.{year}" for day in range(1, num_days + 1)]
        month_df = pd.DataFrame({
            "Дата": days,
            "Сумма": [0] * num_days,
            "Тип траты": ["—"] * num_days,
            "Описание": ["—"] * num_days
        })

    if expense_type == "трата":
        amount = -abs(amount)
    
    new_row = pd.DataFrame({
        "Дата": [f"{day:02}.{month:02}.{year}"],
        "Сумма": [amount],
        "Тип траты": [expense_type],
        "Описание": [description]
    })
    month_df = pd.concat([month_df, new_row], ignore_index=True)

    month_df = month_df[month_df["Дата"] != "Итого"]
    total_amount = month_df["Сумма"].sum()

    total_row = pd.DataFrame({
        "Дата": ["Итого"],
        "Сумма": [total_amount],
        "Тип траты": [""],
        "Описание": [""]
    })
    month_df = pd.concat([month_df, total_row], ignore_index=True)

    month_df.sort_values(by="Дата", inplace=True, ignore_index=True)
    df[month_name] = month_df

    try:
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for sheet_name, sheet_df in df.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print("Данные успешно записаны в Excel.")
    except Exception as e:
        print(f"Ошибка при записи в Excel: {e}")

    history.append((year, month, day, amount, expense_type, description))

# Новая функция для отображения итога за текущий месяц
def show_monthly_total(message):
    today = datetime.now()
    month_name = f"{calendar.month_name[today.month]}"
    
    if os.path.exists(file_name):
        df = pd.read_excel(file_name, sheet_name=None)
        if month_name in df:
            month_df = df[month_name]
            total_row = month_df[month_df["Дата"] == "Итого"]
            if not total_row.empty:
                total_amount = total_row["Сумма"].values[0]
                bot.send_message(
                    message.chat.id,
                    f"Конечно! Вот твой итог за этот месяц: **{total_amount} злотых**", parse_mode="Markdown"
                )
            else:
                bot.send_message(message.chat.id, "Итог за этот месяц еще не рассчитан.")
        else:
            bot.send_message(message.chat.id, f"Лист за месяц {month_name} не найден.")
    else:
        bot.send_message(message.chat.id, "Файл с расходами не найден.")

# Обработка голосовых сообщений
@bot.message_handler(content_types=['voice'])
def handle_voice(message):
    print("Получено голосовое сообщение.")
    file_info = bot.get_file(message.voice.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    with open("voice.ogg", 'wb') as new_file:
        new_file.write(downloaded_file)
    print("Голосовое сообщение сохранено как voice.ogg")

    sound = AudioSegment.from_ogg("voice.ogg")
    sound.export("voice.wav", format="wav")
    print("Файл voice.ogg преобразован в voice.wav для распознавания.")

    recognizer = sr.Recognizer()
    with sr.AudioFile("voice.wav") as source:
        audio = recognizer.record(source)
    try:
        text = recognizer.recognize_google(audio, language="ru-RU").lower().strip()
        print(f"Распознанный текст: {text}")
        bot.reply_to(message, f"Распознанный текст: {text}")

        if "итог" in text:
            show_monthly_total(message)
        elif "убери два прошлых" in text:
            remove_last_entries(2, message)
        elif "убери" in text:
            remove_last_entries(1, message)
        else:
            parse_and_save_expense(text, message)

    except sr.UnknownValueError:
        print("Не удалось распознать голосовое сообщение.")
        bot.reply_to(message, "Извините, не могу распознать голос.")
    except sr.RequestError:
        print("Сервис распознавания временно недоступен.")
        bot.reply_to(message, "Сервис распознавания временно недоступен.")

# Функция для очистки текста от слов-паразитов
def remove_stop_words(text):
    text_lower = text.lower().strip()
    for word in stop_words:
        if text_lower.startswith(word + " "):
            text = text[len(word):].strip()
            break
    return text

# Функция для извлечения информации и обновления файла Excel
def parse_and_save_expense(text, message):
    try:
        print("Анализ текста для извлечения информации о расходах...")

        text = remove_stop_words(text)
        print(f"Текст после удаления слов-паразитов: {text}")

        expense_type = "трата" if "потратил" in text else "заработок"
        print(f"Определён тип траты: {expense_type}")

        zloty = 0
        kopeck = 0

        zloty_pattern = re.search(r"(\d+)\s*(злотых|злотый|злотые|злот)", text)
        if zloty_pattern:
            zloty = int(zloty_pattern.group(1))
            print(f"Извлечённая сумма в злотых: {zloty}")
            text = re.sub(r"\b\d+\s*(злотых|злотый|злотые|злот)\b", "", text).strip()

        kopeck_pattern = re.search(r"(\d+)\s*(коп|копеек)", text)
        if kopeck_pattern:
            kopeck = int(kopeck_pattern.group(1))
            print(f"Извлечённая сумма в копейках: {kopeck}")
            text = re.sub(r"\b\d+\s*(коп|копеек)\b", "", text).strip()

        amount = zloty + kopeck / 100
        print(f"Общая сумма: {amount}")

        description = re.sub(r"(\d+\s*злотых|заработал|доход|прибыль|трата|потратил)", "", text, flags=re.IGNORECASE).strip()
        print(f"Извлечённое описание: {description}")

        today = datetime.now()
        year, month, day = today.year, today.month, today.day

        update_monthly_expenses(year, month, day, amount, expense_type, description)

        bot.send_message(
            message.chat.id,
            f"Готово ✅\nСумма: {amount}\nТип траты: {expense_type}\nОписание: {description}\nДата: {today.strftime('%Y-%m-%d')}"
        )
        print("Сообщение подтверждения отправлено пользователю.")

    except Exception as e:
        print(f"Ошибка: {e}")
        bot.send_message(message.chat.id, "Произошла ошибка при обработке сообщения.")

# Функция для удаления последних записей из Excel
def remove_last_entries(count, message):
    global history

    if len(history) < count:
        bot.send_message(message.chat.id, "Недостаточно записей для удаления.")
        return

    df = pd.read_excel(file_name, sheet_name=None)
    month_name = f"{calendar.month_name[datetime.now().month]}"

    if month_name in df:
        month_df = df[month_name]
        
        for _ in range(count):
            year, month, day, amount, expense_type, description = history.pop()
            condition = (month_df['Дата'] == f"{day:02}.{month:02}.{year}") & (month_df['Сумма'] == amount)
            month_df = month_df[~condition]
        
        total_amount = month_df[month_df['Дата'] != "Итого"]["Сумма"].sum()
        
        if (month_df['Дата'] == "Итого").any():
            month_df.loc[month_df['Дата'] == "Итого", 'Сумма'] = total_amount
        else:
            total_row = pd.DataFrame({"Дата": ["Итого"], "Сумма": [total_amount], "Тип траты": [""], "Описание": [""]})
            month_df = pd.concat([month_df, total_row], ignore_index=True)

        df[month_name] = month_df

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for sheet_name, sheet_df in df.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        bot.send_message(message.chat.id, f"Удалены последние {count} записи и пересчитан итог.")
        print(f"Удалены последние {count} записи и пересчитан итог.")
    else:
        bot.send_message(message.chat.id, "Ошибка: лист текущего месяца не найден.")

print("Бот запущен и готов к работе.")
bot.polling()
