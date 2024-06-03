import speech_recognition as sr
import pyttsx3
import pyautogui
import win32com.client

# Инициализация голосового движка
engine = pyttsx3.init()

# Настройка голоса
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)  # Выбор голоса (0 - мужской, 1 - женский)

# Функция для преобразования текста в речь
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Функция для распознавания речи
def recognize_speech():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Слушаю...")
        audio = r.listen(source)
    try:
        text = r.recognize_google(audio, language="ru-RU")
        print("Вы сказали: " + text)
        return text
    except sr.UnknownValueError:
        print("Извините, я не расслышал.")
        return None
    except sr.RequestError as e:
        print("Ошибка сервиса распознавания речи; {0}".format(e))
        return None

# Функция для работы с документами
def work_with_documents(command):
    # Инициализация Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True

    if "открыть" in command:
        # Открытие нового документа
        word.Documents.Add()
        speak("Документ Word открыт.")

    elif "сохранить" in command:
        # Сохранение документа
        doc = word.ActiveDocument
        doc.SaveAs("document.docx")
        speak("Документ сохранен.")

    elif "закрыть" in command:
        # Закрытие документа
        doc = word.ActiveDocument
        doc.Close()
        word.Quit()
        speak("Документ закрыт.")

    elif "форматировать" in command:
        # Форматирование текста
        doc = word.ActiveDocument
        selection = word.Selection
        selection.Font.Bold = True
        selection.Font.Size = 14
        selection.ParagraphFormat.Alignment = 1  # Выравнивание по центру
        speak("Текст отформатирован.")

    elif "вставить" in command:
        # Вставка текста из голосового ввода
        text = recognize_speech()
        if text:
            doc = word.ActiveDocument
            selection = word.Selection
            selection.TypeText(text)
            speak("Текст вставлен.")

    else:
        speak("Извините, я не понял вашу команду.")

# Основной цикл
while True:
    text = recognize_speech()
    if text:
        work_with_documents(text)
#  # Инициализация Word
#     word = win32com.client.Dispatch("Word.Application")
#     word.Visible = True

#     # Получение активного документа
#     doc = word.ActiveDocument
#     selection = word.Selection

#     if "выделить" in command:
#         # Выделение текста
#         start_pos = selection.Start
#         end_pos = selection.End
#         text = selection.Text
#         speak(f"Выделен текст: {text}")

#     elif "копировать" in command:
#         # Копирование выделенного текста
#         selection.Copy()
#         speak("Текст скопирован в буфер обмена.")

#     elif "вставить" in command:
#         # Вставка текста из буфера обмена
#         selection.Paste()
#         speak("Текст вставлен.")

#     elif "удалить" in command:
#         # Удаление выделенного текста
#         selection.Delete()
#         speak("Текст удален.")

#     elif "заменить" in command:
#         # Замена текста
#         old_text = selection.Text
#         new_text = recognize_speech()
#         if new_text:
#             selection.TypeText(new_text)
#             speak(f"Текст заменен. Было: {old_text}, стало: {new_text}")

#     elif "форматировать" in command:
#         # Форматирование текста
#         selection.Font.Bold = True
#         selection.Font.Size = 14
#         selection.ParagraphFormat.Alignment = 1  # Выравнивание по центру
#         speak("Текст отформатирован.")

#     else:
#         speak("Извините, я не понял вашу команду.")
#№2
# # Текст для обработки
# text = "This is an example sentence. Replace this word with another one."

# # Функция для замены слова в тексте
# def replace_word(text, old_word, new_word):
#     return text.replace(old_word, new_word)

# # Функция для выделения предложения в тексте
# def highlight_sentence(text, sentence):
#     sentences = text.split(". ")
#     for i, s in enumerate(sentences):
#         if sentence in s:
#             return f" Sentence {i+1}: **{s}**"
#     return "Sentence not found"

# try:
#         voice_command = r.recognize_google(audio, language="en-US")
#         print("You said:", voice_command)

#         # Замена слова
#         if "replace" in voice_command:
#             old_word = voice_command.split("replace ").split(" with")
#             new_word = voice_command.split(" with ")
#             text = replace_word(text, old_word, new_word)
#             print("Text after replacement:", text)

#         # Выделение предложения
#         elif "highlight" in voice_command:
#             sentence = voice_command.split("highlight ")
#             highlighted_sentence = highlight_sentence(text, sentence)
#             print("Highlighted sentence:", highlighted_sentence)

#         # Воспроизведение текста
#         engine.say(text)
#         engine.runAndWait()

# except sr.UnknownValueError:
#     print("Sorry, I didn't understand")
# В этом коде мы используем библиотеку speech_recognition для распознавания голоса и pyttsx3 для синтеза речи.
# В главном цикле мы слушаем голос, распознаем его и выполняем соответствующие действия:
# Если в голосовой команде содержится слово "replace", мы заменяем слово в тексте.
# Если в голосовой команде содержится слово "highlight", мы выделяем предложение в тексте.
# В любом случае, мы воспроизводим текст с помощью речевого синтезатора.

