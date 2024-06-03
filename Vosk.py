import os
import wave
import vosk
import json
import pyttsx3
import pyautogui
import win32com.client
import datetime
import webbrowser
import pyaudio
import subprocess

# Функция для преобразования текста в речь
def speak(text):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)  # Выбор голоса (0 - мужской, 1 - женский)
    engine.say(text)
    engine.runAndWait()

model = vosk.Model("model_small")
# Функция для распознавания речи
def recognize_speech():
    # Инициализируем модель Vosk для русского языка
    # model = vosk.Model("model_small")
    
    # Создаем объект распознавания речи
    rec = vosk.KaldiRecognizer(model, 16000)
    
    # Открываем микрофон
    p = pyaudio.PyAudio()
    stream = p.open(format=pyaudio.paInt16,
                    channels=1,
                    rate=16000,
                    input=True,
                    frames_per_buffer=1024)
    
    # Распознаем речь
    while True:
        data = stream.read(4000)
        if len(data) == 0:
            break
        if rec.AcceptWaveform(data):
            result = json.loads(rec.Result())
            text = result["text"]
            print("Вы сказали: " + text)
            # if text:
            #     work_with_documents(text)
            if "стоп" in text.lower():
                break
            return text

    
    # Закрываем микрофон
    stream.stop_stream()
    stream.close()
    p.terminate()
    return None
def insert_document():
    # Распознаем речь
    text = recognize_speech()
    if text:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.ActiveDocument
        selection = word.Selection
        selection.TypeText(text)
        speak("Текст вставлен.")
# Функция для работы с документами
def work_with_documents(command):
   
    word = win32com.client.Dispatch("Word.Application")
    if "время" in command:
        now = datetime.datetime.now()
        speak("Сейчас " + str(now.hour) + ":" + str(now.minute))

    elif "открыть документ"==command:
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
        speak("Скажите текст для вставки")
        insert_document()

    elif "выделить" in command:
        # Выделение текста
        pyautogui.hotkey('ctrl','a')
        speak("текст выделен")

    elif "копировать" in command:
        # Копирование выделенного 
        pyautogui.hotkey('ctrl','c')
        speak("скопировано")

    elif "буфер" in command:
        # Вставка текста из буфера обмена
        pyautogui.hotkey('ctrl','v')
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
        speak("Открываю проводник")
        pyautogui.hotkey('win','e')

    elif "музыка" in command:
        webbrowser.open_new_tab('https://music.yandex.ru/home')
    else:
        speak("Извините, я не понял вашу команду.")
        return True
while True:
    text=recognize_speech()
    if text:
        work_with_documents(text)
    