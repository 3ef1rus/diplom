import sys
import speech_recognition as sr
import pyttsx3
import os
import nltk
import subprocess
import webbrowser
import random
import keyboard
import ctypes
import time
import threading
import pyautogui


from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QApplication,QDialog, QComboBox,QLineEdit
from PyQt5.QtCore import QTimer,pyqtSignal
from PyQt5 import uic
from pathlib import Path
nltk.download('punkt')
from nltk.tokenize import word_tokenize
nltk.download('averaged_perceptron_tagger')
from nltk import pos_tag
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume



os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "путь_к_ключу_доступа.json"

class MyCustomError(Exception):
    pass



def killProg():
    time.sleep(2)
    ctypes.windll.kernel32.ExitProcess(0)

def searchInBrows(text=""):
    if text=="":
        webbrowser.open('https://www.google.com')
    else :
        url = f"https://www.google.com/search?q={text}"
        webbrowser.open(url)
        
def searchVideo(text=""):
    if text=="":
        webbrowser.open('https://www.youtube.com')
    else :
        url = f"https://www.youtube.com/search?q={text}"
        webbrowser.open(url)

def restart_program():
    # """Перезагружает текущую программу."""
    script_path = os.path.realpath(__file__)
    subprocess.call(['python', script_path])
    
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

def choseSayOK():
    my_list = ["Хорошо сейчас сделаю", "Сейчас", "Хорошо", "Угу", "Секунду","Готово","В процессе","Сделаю","Уже делаю"]
    random_element = random.choice(my_list)
    return (random_element)

def choseSay():
    my_list = ["Возможно вы попросили то что я еще не умею", "Я такое не умею", "Я вас не поняла", "Попробуйте сформулировать по другому", "Очень хотелось бы помочь но я вас не поняла","Не понятно"]
    random_element = random.choice(my_list)
    return (random_element)

def freeze_button(button, seconds):
    button.setEnabled(False)  # Замораживаем кнопку
    QTimer.singleShot(seconds * 1000, lambda: button.setEnabled(True))

class ChildWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi("childWindow.ui", self)
        
    def get_combobox_value(self):
        return self.comboBox.currentText()

class VoiceAssistantApp(QtWidgets.QMainWindow):
    
    def __init__(self):
        super(VoiceAssistantApp, self).__init__()
        uic.loadUi("speeGUI.ui", self)  
        self.show()
        self.child_window = ChildWindow()
        # создаем объект для распознавания речи
        
        self.r = sr.Recognizer()
        self.i = 0
        # создаем объект для синтеза речи
        self.engine = pyttsx3.init()
        self.voices = self.engine.getProperty('voices')
        self.lineedit_text = threading.Thread(target=self.get_lineedit_text)
        # Подключение сигналов кнопок к соответствующим методам
        self.lineEnter.returnPressed.connect(self.get_lineedit_text)
        self.btn_start.clicked.connect(self.start_listening)
        self.btn_stop.clicked.connect(self.stop_listening)
        self.btn_tool.clicked.connect(self.open_child_window)
        self.btn_Enter.clicked.connect(self.get_lineedit_text)
        
    def check_thread(self,thread):
            if self.thread.is_alive():
                self.thread.join()
    
    def get_lineedit_text(self):
        text = self.lineEnter.text()
        self.lineEnter.clear()
        print("Текст из LineEdit:", text)
        self.append_text(f"Вы написали: {text}")
        self.process_user_input(text)
                 
    def access_combobox(self):
        
        combobox_value = self.child_window.get_combobox_value()
        return combobox_value
    
    def open_child_window(self):
        
        child_window = ChildWindow()
        child_window.exec_()
        
    def closeEvent(self, event):
        # Вызывается при нажатии кнопки закрытия окна
        print("Close button clicked")
        killProg()
        # Вызов базовой реализации события closeEvent
        super().closeEvent(event)
    
    def start_listening(self):
        
        self.textChat.clear()
        self.btn_start.setEnabled(False)
        self.listening=True
        self.btn_stop.setEnabled(True)
        self.start_listening_thread = threading.Thread(target=self.continuous_listening)
        self.start_listening_thread.start()

    def stop_listening(self):   
        
        self.stop_listening_thread = threading.Thread(target=self.stop_list_trh)
        self.stop_listening_thread.start()
        
    def stop_list_trh(self):
        
        self.btn_stop.setEnabled(False)
        self.listening=False
        self.start_listening_thread.join()
        self.btn_start.setEnabled(True)
         
    def append_text(self, text):
        
        self.textChat.insertPlainText(text + "\n")
        self.textChat.ensureCursorVisible()
        # self.check_thread(self.le)

    def continuous_listening(self):
        
        while self.listening:
            self.listen_for_speech()

    def speak(self, text):
        
        self.append_text("Spee: " + text)
        # self.engine.setProperty('voice', self.voices[0].id)
        self.engine.say(text)
        self.engine.runAndWait()

    def listen_for_speech(self):
        if self.listening:
            with sr.Microphone() as source:
                self.append_text("Spee: " +"Говорите...")
                print("Говорите...")
                self.r.adjust_for_ambient_noise(source, duration=1)
                self.audio = self.r.listen(source)
                try:
                    if self.listening:  # Проверка флага на прослушивание 
                        text = self.r.recognize_google (self.audio, language="ru-RU")
                        self.append_text(f"Вы сказали: {text}")
                        print(f"Вы сказали: {text}")
                        self.process_user_input(text)
                except sr.UnknownValueError:
                    self.append_text("Spee: " +"Речь не распознана")
                    print("Речь не распознана")
                    if self.listening:  # Проверка флага на прослушивание
                        self.listen_for_speech()
                except sr.RequestError as e:
                    self.append_text("Spee: " +f"Ошибка: {e}")
                    if self.listening:  # Проверка флага на прослушивание
                        self.listen_for_speech()

    def process_user_input(self, text):
            
        tokens = word_tokenize(text.lower())
        tags = pos_tag(tokens)
        print(tags)
        # проверяем условия для ответа на запросы пользователя
        if "привет" in text.lower():
            self.speak("Привет, как я могу вам помочь?")
            
        elif "как дела" in text.lower():
                self.speak("У меня все хорошо, спасибо за интерес.")
                
        elif "спасибо" in text.lower():
            self.speak("Всегда пожалуйста.")
            
        elif "создай блокнот на рабочем столе" in text.lower():
            self.speak(choseSayOK())
            if "с названием" in text.lower():           
                for word, tag in reversed(tags):
                        if tag == 'NN':
                            x=word
                            createFile("txt",x)
                            break
            else:createFile("txt")
        
        elif "создай word на рабочем столе" in text.lower():
            self.speak(choseSayOK())
            if "с названием" in text.lower():           
                for word, tag in reversed(tags):
                        if tag == 'NN':
                            x=word
                            createFile("doc",x)
                            break
            else: createFile("doc")
            
        elif "создай excel на рабочем столе" in text.lower():
            self.speak(choseSayOK())

            if "с названием" in text.lower():           
                for word, tag in reversed(tags):
                        if tag == 'NN':
                            x=word
                            createFile("xlsx",x)
                            break
            else:createFile("xlsx")    
        
        elif "создай папку на рабочем столе" in text.lower():
            self.speak(choseSayOK())
            if "с названием" in text.lower():           
                for word, tag in tags:
                        if tag == 'NN':
                            x=word
                            createFolder(x)
            else:createFolder("")                
                    
        elif "создай презентацию на рабочем столе" in text.lower():
            self.speak(choseSayOK())
            if "с названием" in text.lower():           
                for word, tag in reversed(tags):
                        if tag == 'NN':
                            x=word
                            createFile("pptx",x)
                            break
            else:createFile("pptx")  
            
        elif "открой google" in text.lower():
            self.speak(choseSayOK())
            if "и найди" in text.lower():           
                string=text.lower()
                idx=string.find("найди")
                second_half = string[idx + len("найди"):].strip()
                searchInBrows(second_half)
            else:searchInBrows()    
            
        elif "открой youtube" in text.lower():
            self.speak(choseSayOK())
            if "и найди" in text.lower():           
                string=text.lower()
                idx=string.find("найди")
                second_half = string[idx + len("найди"):].strip()
                searchVideo(second_half)
            else:searchVideo()                   
            
        elif "открой калькулятор" in text.lower():
            self.speak(choseSayOK())
            subprocess.Popen('calc.exe')     
            
        elif "сделай тише" in text.lower():
            for word, tag in tags:
                if tag == 'CD':
                    changeVolumeMIN(int(word)/100)
            self.speak(choseSayOK())
            
        elif "сделай потише" in text.lower():
            for word, tag in tags:
                if tag == 'CD':
                    changeVolumeMIN(int(word)/100)
            self.speak(choseSayOK())    
                
        elif "сделай громче" in text.lower():
            for word, tag in tags:
                if tag == 'CD':
                    changeVolumeMAX(int(word)/100)
            self.speak(choseSayOK())

        elif "сделай погромче" in text.lower():
            for word, tag in tags:
                if tag == 'CD':
                    changeVolumeMAX(int(word)/100)
            self.speak(choseSayOK())
            
        elif "сделай звук на максимум" in text.lower():
            self.speak(choseSayOK())
            changeVolumeM()
            
        elif "сверни все окна" in text.lower():
            self.speak(choseSayOK())
            pyautogui.keyDown('win')
            pyautogui.press('d')
            pyautogui.keyUp('win')
            
        elif "выключи звук" in text.lower():
            self.speak(choseSayOK())
            MuteVolume()
        
        elif "перезагрузись" in text.lower():
            self.speak(choseSayOK())
            restart_program()
        
        elif "открой bluetooth" in text.lower():
            self.speak(choseSayOK())
            subprocess.Popen('explorer.exe shell:::{28803F59-3A75-4058-995F-4EE5503B023C}')
                    
        elif "поменяй язык на клавиатуре" in text.lower():
            self.speak(choseSayOK())
            keyboard.press_and_release('alt+shift')   
                    
        elif "стоп" in text.lower():
            self.speak("До свидания.")
            killProg()
        
        else : self.speak(choseSay())
            
        
    


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VoiceAssistantApp()
    window.show()
    sys.exit(app.exec_())
