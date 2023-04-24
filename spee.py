import speech_recognition as sr
import pyttsx3
import os
import nltk
import pyautogui
import subprocess
from pathlib import Path
nltk.download('punkt')
from nltk.tokenize import word_tokenize
nltk.download('averaged_perceptron_tagger')
from nltk import pos_tag
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# создаем объект для распознавания речи
r = sr.Recognizer()
i=0
# создаем объект для синтеза речи
engine = pyttsx3.init()

class MyCustomError(Exception):
    pass

# функция для преобразования текста в речь
def speak(text):
    engine.say(text)
    engine.runAndWait()

def changeVolumeMIN(x):
    # Получение всех устройств воспроизведения звука
    if x>1 or x<0 : x=0.1
    sessions = AudioUtilities.GetAllSessions()
    # Цикл по всем сессиям и увеличение громкости на 10%
    try:
        for session in sessions:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            current_volume = volume.GetMasterVolume()
            volume.SetMasterVolume(current_volume - x, None)
    except: print("Ошибка")

def changeVolumeMAX(x):
    # Получение всех устройств воспроизведения звука
    if x>1 or x<0 : x=0.1
    sessions = AudioUtilities.GetAllSessions()

    # Цикл по всем сессиям и увеличение громкости на 10%
    try:
        for session in sessions:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            current_volume = volume.GetMasterVolume()
            volume.SetMasterVolume(current_volume + x, None)
    except: print("Ошибка")
    
def changeVolumeM():
    # Получение всех устройств воспроизведения звука
    sessions = AudioUtilities.GetAllSessions()

    # Цикл по всем сессиям и увеличение громкости на 10%
    try:
        for session in sessions:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMasterVolume(1.0, None)
    except: print("Ошибка")
            
def MuteVolume():
    # Получение всех устройств воспроизведения звука
    sessions = AudioUtilities.GetAllSessions()

    # Цикл по всем сессиям и увеличение громкости на 10%
    try:
        for session in sessions:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMasterVolume(0, None)
    except: print("Ошибка")

def createFile(x,name="file"):
     # Получаем путь к рабочему столу пользователя
        desktop_path = Path.home() / "Desktop"
        # Указываем имя файла и его расширение
        if x=="doc":
            name=f'{name}.doc'
            file_path = desktop_path / name
        elif x=="txt":
            name=f'{name}.txt'
            file_path = desktop_path / name
        elif x=="pptx":
            name=f'{name}.pptx'
            file_path = desktop_path / name
        elif x=="xlsx":
            name=f'{name}.xlsx'
            file_path = desktop_path / name
            

        # Создаем пустой файл на рабочем столе
        try:
            if file_path.exists():
              raise MyCustomError("Файл с таким названием уже существует")
            else: file_path.touch()          
        except MyCustomError:    
            i = 1
            while True:
                new_file_path = desktop_path / f'{i}_{name}'
                try:
                    if new_file_path.exists():
                        raise MyCustomError("Файл с таким названием уже существует")
                    else: new_file_path.touch()  
                    file_path = new_file_path
                    break
                except MyCustomError:
                    i += 1

def createFolder(x):
    # Путь к рабочему столу
    global i
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    # Название новой папки
    folder_name="Новая папка"
    
    # Полный путь к новой папке
    folder_path = os.path.join(desktop_path, folder_name)

    # Создание новой папки
    if x=="":
        try:
            os.mkdir(folder_path)           
        except FileExistsError:    
            i = 1
            while True:
                new_folder_name = f"{folder_name}_{i}"
                new_folder_path = os.path.join(desktop_path, new_folder_name)
                try:
                    os.mkdir(new_folder_path)
                    folder_name = new_folder_name
                    folder_path = new_folder_path
                    break
                except FileExistsError:
                    i += 1
    
    else:
        folder_name = x
        folder_path = os.path.join(desktop_path, folder_name)
        try:
            os.mkdir(folder_path)           
        except FileExistsError:    
            i = 1
            while True:
                new_folder_name = f"{folder_name}_{i}"
                new_folder_path = os.path.join(desktop_path, new_folder_name)
                try:
                    os.mkdir(new_folder_path)
                    folder_name = new_folder_name
                    folder_path = new_folder_path
                    break
                except FileExistsError:
                    i += 1                
    
# функция для распознавания речи
def recognize_speech():
    with sr.Microphone() as source:
        print("Говорите...")
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)

    try:
        text = r.recognize_google(audio, language="ru-RU")
        print(f"Вы сказали: {text}")
        return text
    except sr.UnknownValueError:
        print("Речь не распознана")
        return ""
    except sr.RequestError as e:
        print(f"Ошибка: {e}")
        return ""

# основной цикл работы голосового помощника
while True:
    # распознаем речь пользователя
    text = recognize_speech()
    tokens = word_tokenize(text.lower())
    tags = pos_tag(tokens)
    print(tags)
    # проверяем условия для ответа на запросы пользователя
    if "привет" in text.lower():
        speak("Привет, как я могу вам помочь?")
        
    elif "как дела" in text.lower():
        speak("У меня все хорошо, спасибо за интерес.")
        
    elif "спасибо" in text.lower():
        speak("Всегда пожалуйста.")
        
    elif "создай блокнот на рабочем столе" in text.lower():
        speak("Хорошо сейчас сделаю.")
        if "с названием" in text.lower():           
            for word, tag in reversed(tags):
                    if tag == 'NN':
                        x=word
                        createFile("txt",x)
        createFile("txt")
    
    elif "создай word на рабочем столе" in text.lower():
        speak("Хорошо сейчас сделаю.")
        if "с названием" in text.lower():           
            for word, tag in reversed(tags):
                    if tag == 'NN':
                        x=word
                        createFile("doc",x)
                        break
        else: createFile("doc")
        
    elif "создай excel на рабочем столе" in text.lower():
        speak("Хорошо сейчас сделаю.")

        if "с названием" in text.lower():           
            for word, tag in reversed(tags):
                    if tag == 'NN':
                        x=word
                        createFile("xlsx",x)
        createFile("xlsx")    
    
    elif "создай папку на рабочем столе" in text.lower():
        speak("Хорошо сейчас сделаю.")
        if "с названием" in text.lower():           
            for word, tag in tags:
                    if tag == 'NN':
                        x=word
                        createFolder(x)
        else:createFolder("")                
                
              
    elif "создай презентацию на рабочем столе" in text.lower():
        speak("Хорошо сейчас сделаю.")
        if "с названием" in text.lower():           
            for word, tag in reversed(tags):
                    if tag == 'NN':
                        x=word
                        createFile("pptx",x)
        createFile("pptx")  
        
    elif "открой калькулятор" in text.lower():
        speak("Хорошо сейчас сделаю.")
        subprocess.Popen('calc.exe')     
        
    elif "сделай тише" in text.lower():
        for word, tag in tags:
            if tag == 'CD':
                changeVolumeMIN(int(word)/100)
        speak("Хорошо сейчас сделаю.")
        
    elif "сделай потише" in text.lower():
        for word, tag in tags:
            if tag == 'CD':
                changeVolumeMIN(int(word)/100)
        speak("Хорошо сейчас сделаю.")    
              
    elif "сделай громче" in text.lower():
        for word, tag in tags:
            if tag == 'CD':
                changeVolumeMAX(int(word)/100)
        speak("Хорошо сейчас сделаю.")

    elif "сделай погромче" in text.lower():
        for word, tag in tags:
            if tag == 'CD':
                changeVolumeMAX(int(word)/100)
        speak("Хорошо сейчас сделаю.")
        
    elif "сделай звук на максимум" in text.lower():
        
        speak("Хорошо сейчас сделаю.")
        changeVolumeM()
        
    elif "сверни все окна" in text.lower():
        
        speak("Хорошо сейчас сделаю.")
        pyautogui.minimizeAllWindows()
        
    elif "выключи звук" in text.lower():
        
        speak("Хорошо сейчас сделаю.")
        MuteVolume()
             
          
    elif "стоп" in text.lower():
        speak("До свидания.")
        break
