import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta
import pandas as pd
import openpyxl

class WorkTimeTracker:
    def __init__(self, root):
        self.root = root
        self.root.title("Work Time Tracker")

        # Загрузка ФИО из файла
        self.names = self.load_names_from_file("FIO.txt")

        # Создание выпадающего списка
        self.name_var = tk.StringVar()
        self.name_dropdown = ttk.Combobox(root, textvariable=self.name_var, values=self.names)
        self.name_dropdown.grid(row=0, column=0, padx=10, pady=10)

        # Кнопки старт и стоп
        self.start_button = tk.Button(root, text="Старт", command=self.start_timer)
        self.start_button.grid(row=0, column=1, padx=10, pady=10)

        self.stop_button = tk.Button(root, text="Стоп", command=self.stop_timer, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=2, padx=10, pady=10)

        # Таймер и временные метки
        self.start_time = None
        self.start_button_time = None
        self.stop_button_time = None

    def load_names_from_file(self, filename):
        with open(filename, 'r', encoding='utf-8') as file:
            names = file.read().splitlines()
        return names

    def start_timer(self):
        self.start_time = datetime.now()
        self.start_button_time = self.start_time.strftime("%H:%M")
        self.start_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)

    def stop_timer(self):
        if self.start_time:
            end_time = datetime.now()
            elapsed_time = end_time - self.start_time
            fio = self.name_var.get().split(' ')
            last_name, first_name, patronymic = fio if len(fio) == 3 else ('', '', '')
            date = datetime.now().strftime("%d-%m-%Y")
            
            # Исправление ошибки в преобразовании времени
            hours, remainder = divmod(elapsed_time.seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            
            time_worked = '{:02}:{:02}'.format(int(hours), int(minutes))
            stop_button_time = end_time.strftime("%H:%M")

            # Определение значения для нового столбца
            work_status, overtime = self.calculate_work_status(self.start_button_time, stop_button_time)

            # Запись данных в Excel
            self.save_to_excel(last_name, first_name, patronymic, date, time_worked, self.start_button_time, stop_button_time, work_status, overtime)

            self.start_button.configure(state=tk.NORMAL)
            self.stop_button.configure(state=tk.DISABLED)
            self.start_time = None
            self.start_button_time = None

    def calculate_work_status(self, start_time, stop_time):
        start_datetime = datetime.strptime(start_time, "%H:%M")
        stop_datetime = datetime.strptime(stop_time, "%H:%M")

        work_status = ""
        overtime = "00:00"

        if start_datetime < datetime.strptime("09:00", "%H:%M") and stop_datetime > datetime.strptime("17:00", "%H:%M"):
            work_status = "Отработано"
        elif start_datetime > datetime.strptime("09:00", "%H:%M") and stop_datetime > datetime.strptime("17:00", "%H:%M"):
            work_status = "Опоздание"
            overtime = '{:02}:{:02}'.format(*divmod((start_datetime - datetime.strptime("09:00", "%H:%M")).seconds, 60))
        elif start_datetime < datetime.strptime("09:00", "%H:%M") and stop_datetime > datetime.strptime("17:10", "%H:%M"):
            work_status = "Переработка"
            overtime = '{:02}:{:02}'.format(*divmod((stop_datetime - datetime.strptime("17:00", "%H:%M")).seconds, 60))

        return work_status, overtime

    def save_to_excel(self, last_name, first_name, patronymic, date, time_worked, start_button_time, stop_button_time, work_status, overtime):
        df = pd.DataFrame({
            "Фамилия": [last_name],
            "Имя": [first_name],
            "Отчество": [patronymic],
            "Дата": [date],
            "Время отработанное": [time_worked],
            "Время старта": [start_button_time],
            "Время стопа": [stop_button_time],
            "Статус работы": [work_status],
            "Переработка": [overtime]
        })

        try:
            existing_data = pd.read_excel("work_time_data.xlsx")
            df = pd.concat([existing_data, df], ignore_index=True)
        except FileNotFoundError:
            pass

        # Запись данных в Excel
        with pd.ExcelWriter("work_time_data.xlsx", engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1', startrow=len(existing_data) + 1, header=False)

if __name__ == "__main__":
    root = tk.Tk()
    app = WorkTimeTracker(root)
    root.mainloop()
