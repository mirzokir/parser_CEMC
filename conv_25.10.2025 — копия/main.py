#!/usr/bin/env python3
"""
Главный лаунчер для четырёх парсеров.
Файлы (должны быть в той же папке):
 1 — rrl_incoming_parser.py
 2 — rrl_outgoing_parser.py
 3 — спс_incoming_parser.py
 4 — спс_outgoing_parser.py

После запуска меню позволяет выбирать скрипт, запускает его тем же интерпретатором
(sys.executable). По завершении возвращается в меню.
"""
import subprocess
import sys
import os

FILES = {
    "1": "rrl_incoming_parser.py",
    "2": "rrl_outgoing_parser.py",
    "3": "спс_incoming_parser.py",
    "4": "спс_outgoing_parser.py",
}

MENU = '''
Выберите действие:
 1 — RRL Incoming Parser
 2 — RRL Outgoing Parser
 3 — СПС Incoming Parser
 4 — СПС Outgoing Parser
 0 — Выход
'''


def run_script(path):
    """Запускает скрипт через текущий интерпретатор Python и возвращает код завершения."""
    try:
        # Включаем вывод дочернего процесса в тот же терминал, чтобы скрипты могли взаимодействовать с пользователем
        completed = subprocess.run([sys.executable, path])
        return completed.returncode
    except FileNotFoundError:
        print(f"Файл не найден: {path}")
        return -1
    except Exception as e:
        print(f"Ошибка при запуске {path}: {e}")
        return -1


def main():
    # Рабочая директория — та, в которой лежит main.py
    os.chdir(os.path.dirname(os.path.abspath(__file__)) or os.getcwd())

    while True:
        print(MENU)
        choice = input("Введите номер (0-4): ").strip()

        if choice in ("0", "q", "quit", "exit"):
            print("Выход. До встречи!")
            break

        if choice not in FILES:
            print("Неверный ввод — введите 0,1,2,3 или 4.")
            continue

        script_name = FILES[choice]
        script_path = os.path.join(os.getcwd(), script_name)

        if not os.path.isfile(script_path):
            print(f"Скрипт {script_name} не найден в папке: {os.getcwd()}")
            input("Положите файл в эту папку и нажмите Enter, чтобы вернуться в меню...")
            continue

        print(f"\nЗапускаю {script_name}...")
        #print("(Если скрипт спрашивает путь к .txt — просто введите его там)")

        code = run_script(script_path)

        print(f"\nСкрипт {script_name} завершился с кодом: {code}\n")
        input("Нажмите Enter, чтобы вернуться в меню...")


if __name__ == "__main__":
    main()
