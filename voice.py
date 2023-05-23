from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

def set_volume_percent(percent):
    # Получить объект управления громкостью
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    print(current_volume)
    # Установить новую громкость в процентах
    result  = current_volume - (1 * percent) / 100
    volume.SetMasterVolumeLevelScalar(result, None)

# Использование функции для установки громкости в 50 процентов
set_volume_percent(10)
