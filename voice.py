import pyttsx3

def get_available_voice_speeds():
    engine = pyttsx3.init()
    speeds = []

    # Получить текущее значение скорости голоса
    current_speed = engine.getProperty('rate')
    speeds.append(current_speed)

    # Исследовать другие значения скорости голоса
    for speed in range(50, 1000, 50):
        # Установить временную скорость голоса
        engine.setProperty('rate', speed)

        # Проверить, что скорость голоса была установлена
        if engine.getProperty('rate') != current_speed:
            speeds.append(speed)

    return speeds

# Получить список доступных скоростей голоса
available_speeds = get_available_voice_speeds()

# Вывести список скоростей голоса
print("Доступные скорости голоса:")
for speed in available_speeds:
    print(speed)
