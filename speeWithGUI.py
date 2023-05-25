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
import openai

from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
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
    
def changeVolumeMIN(x=10):
    if x>100 or x<0 : x=10
    # Получение всех устройств воспроизведения звука
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    # Установить новую громкость в процентах
    result  = current_volume - (1 * x) / 100

    try:
        
        volume.SetMasterVolumeLevelScalar(result, None)

    except: print("Ошибка")

def changeVolumeMAX(x=10):
    if x>100 or x<0 : x=10
    # Получение всех устройств воспроизведения звука
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    # Установить новую громкость в процентах
    result  = current_volume + (1 * x) / 100

    try:
        
        volume.SetMasterVolumeLevelScalar(result, None)

    except: print("Ошибка")
    
def changeVolumeM():
    # Получение всех устройств воспроизведения звука
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    try:
        
        volume.SetMasterVolumeLevelScalar(1, None)

    except: print("Ошибка")
            
def MuteVolume():
    # Получение всех устройств воспроизведения звука
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    try:
        
        volume.SetMasterVolumeLevelScalar(0, None)

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

def ChatGPTrequest(text):
    openai.api_key = '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'
    
    
    response = openai.Completion.create(
        engine='text-davinci-003',  # Используйте GPT-3.5 модель
        prompt='Ваш текстовый запрос',  # Ваш запрос
        max_tokens=100   
    )
    generated_text = response.choices[0].text.strip()
    return  generated_text

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
    button_clicked = pyqtSignal(str)
    
    def __init__(self,parent=None):
        super().__init__(parent)
        uic.loadUi("childWindow.ui", self)
        self.setWindowTitle("Настройки")
        self.btn_Push.clicked.connect(self.send_info)
    
    def send_info(self):
        Rate=self.get_comboRate_value()
        Volume=self.get_comboVolume_value()
        info=f"{Rate},{Volume}"
        self.signal.emit(info)
    def get_info(self):
        Rate=self.get_comboRate_value()
        self.get_comboVolume_value()
        print(Rate)
        
    def get_comboRate_value(self):
        Rate=self.comboRate.currentText()
        return Rate
    
    def get_comboVolume_value(self):
        
        Volume=self.comboVolume.currentText()
        return Volume
    

        

class VoiceAssistantApp(QtWidgets.QMainWindow):
    
    def __init__(self):
        super(VoiceAssistantApp, self).__init__()
        uic.loadUi("speeGUI.ui", self)  
        self.show()
        self.child_window = ChildWindow()

        self.child_window.button_clicked.connect(self.receive_info)
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
                 
    def receive_info(self,info):
        print("Получена информация:", info)
       
    
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

    def continuous_listening(self):
        
        while self.listening:
            self.listen_for_speech()

    def settingVoice(self):
        volume=self.access_comboVolume()
        rate=self.access_comboRate()
        self.engine.setProperty('volume', volume)
        self.engine.setProperty('rate', rate)

    def speak(self, text):
        
        self.append_text("Spee: " + text)
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
                string=text.lower()
                idx=string.find("названием")
                second_half = string[idx + len("названием"):].strip()
                createFile("txt",second_half)
            else:createFile("txt") 
        
        elif "создай word на рабочем столе" in text.lower():
            self.speak(choseSayOK())
            if "с названием" in text.lower():           
                string=text.lower()
                idx=string.find("названием")
                second_half = string[idx + len("названием"):].strip()
                createFile("doc",second_half)
            else:createFile("doc")  
            
        elif "создай excel на рабочем столе" in text.lower():
            self.speak(choseSayOK())
            if "с названием" in text.lower():           
                string=text.lower()
                idx=string.find("названием")
                second_half = string[idx + len("названием"):].strip()
                createFile("xlsk",second_half)
            else:createFile("xlsk")      
        
        elif "создай папку на рабочем столе" in text.lower():
            self.speak(choseSayOK())
            if "с названием" in text.lower():           
                string=text.lower()
                idx=string.find("названием")
                second_half = string[idx + len("названием"):].strip()
                createFolder(second_half)
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
            
        elif "спроси у помощника" in text.lower():
            self.speak("Данная функция нуждается в спонсоре")
            # self.speak(choseSayOK())         
            # string=text.lower()
            # idx=string.find("помощника")
            # second_half = string[idx + len("помощника"):].strip()
            # self.speak(ChatGPTrequest(second_half))    
            
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
            change=0
            for word, tag in tags:
                if tag == 'CD':
                    changeVolumeMIN(int(word))
                    change+=1
            if change==0:
                changeVolumeMIN()
            self.speak(choseSayOK()) 
            
        elif "сделай потише" in text.lower():
            change=0
            for word, tag in tags:
                if tag == 'CD':
                    changeVolumeMIN(int(word))
                    change+=1
            if change==0:
                changeVolumeMIN()
            self.speak(choseSayOK()) 
                
        elif "сделай громче" in text.lower():
            change=0
            for word, tag in tags:
                if tag == 'CD':
                    changeVolumeMAX(int(word))
                    change+=1
            if change==0:
                changeVolumeMAX()
            self.speak(choseSayOK())

        elif "сделай погромче" in text.lower():
            change=0
            for word, tag in tags:
                if tag == 'CD':
                    changeVolumeMAX(int(word))
                    change+=1
            if change==0:
                changeVolumeMAX()
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
