import speech_recognition as sr
import pyttsx3
import pyautogui
import win32com.client
import datetime
import webbrowser

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
        if "стоп" in text.lower():
            
            return False

        return text
    except sr.UnknownValueError:
        print("Извините, я не расслышал.")
        return None
    except sr.RequestError as e:
        print("Ошибка сервиса распознавания речи; {0}".format(e))
        return None
    
# Функция для работы с документами
def work_with_documents(command):

    word = win32com.client.Dispatch("Word.Application")

    if "время" in command:
        now = datetime.datetime.now()
        speak("Сейчас " + str(now.hour) + ":" + str(now.minute))

    elif "открыть" in command:
        # Открытие нового документа
        
        word.Visible = True
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

    elif "выделить" in command:
        # Выделение текста
        pyautogui.hotkey('ctrl', 'a')
        speak("текст выделен")

    elif "копировать" in command:
        # Копирование выделенного 
        pyautogui.hotkey('ctrl', 'c')
        speak("скопировано")

    elif "буфер" in command:
        # Вставка текста из буфера обмена
        pyautogui.hotkey('ctrl', 'v')
        speak("Вставлен буфер")

    elif "удалить" in command:
        # Удаление выделенного текста
        pyautogui.hotkey('delete')
        speak("текст удален")
    elif "справа"in command:
        pyautogui.hotkey('ctrl','r')
        speak("текст форматирован направо")

    elif "подчеркивание"in command:
        pyautogui.hotkey('ctrl','u')
        speak("подчеркивание")

    elif "ширина"in command:
        pyautogui.hotkey('ctrl','j')
        speak("форматирован по ширине")

    elif "cлева"in command:
        pyautogui.hotkey('ctrl','l')
        speak("форматирован налево")

    elif "курсив"in command:
        pyautogui.hotkey('ctrl','i')
        speak("курсив")

    elif "проводник" in command:
        pyautogui.hotkey('win','e')

    elif "музыка" in command:
        webbrowser.open_new_tab('https://music.yandex.ru/home')
    else:
        speak("Извините, я не понял вашу команду.")
        return True
    pyautogui.hotkey('ctrl','r')
# Основной цикл
while True:
    text = recognize_speech()
    if text:
        work_with_documents(text)
    