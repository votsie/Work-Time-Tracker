import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta
import pandas as pd
import os

class WorkTimerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Work Timer App")
        self.root.geometry("250x150")

        self.fio_list = self.read_fio_from_file("FIO.txt")
        if not self.fio_list:  # Проверка на пустой список ФИО
            self.fio_list = ["Добавьте ФИО в FIO.txt"]

        self.selected_fio = tk.StringVar()
        self.selected_fio.set(self.fio_list[0])

        self.fio_dropdown = ttk.Combobox(self.root, textvariable=self.selected_fio, values=self.fio_list)
        self.fio_dropdown.pack(pady=10)

        self.start_button = tk.Button(self.root, text="Старт", command=self.start_timer)
        self.start_button.pack()

        self.stop_button = tk.Button(self.root, text="Стоп", command=self.stop_timer, state=tk.DISABLED)
        self.stop_button.pack()

        self.timer_label = tk.Label(self.root, text="")
        self.timer_label.pack(pady=10)

        self.start_time = None


    def read_fio_from_file(self, filename):
        try:
            with open(filename, "r", encoding="utf-8") as f:
                return [line.strip() for line in f.readlines()]
        except FileNotFoundError:
            return []  # Возвращаем пустой список, если файл не найден


    def start_timer(self):
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.start_time = datetime.now()
        self.update_timer()

    def stop_timer(self):
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        end_time = datetime.now()
        work_duration = end_time - self.start_time
        fio = self.selected_fio.get()
        if fio != "Добавьте ФИО в FIO.txt":  # Проверка на placeholder ФИО
            self.save_to_excel(fio, self.start_time, end_time, work_duration)
        self.timer_label.config(text=f"Время работы: {work_duration}")
        self.start_time = None  # Сбрасываем start_time после остановки

    def update_timer(self):
        if self.start_time:  # Проверяем, запущено ли время
            elapsed_time = datetime.now() - self.start_time
            self.timer_label.config(text=f"Зафиксированное время: {elapsed_time}")
            self.root.after(1000, self.update_timer)

    def save_to_excel(self, fio, start_time, end_time, work_duration):
        try:
            fio_data = fio.split()
            last_name, first_name, patronymic = fio_data[0], fio_data[1], fio_data[2]
        except IndexError:  # Обработка некорректного формата ФИО
            print("Ошибка: некорректный формат ФИО в FIO.txt")
            return

        start_time_str = start_time.strftime("%Y-%m-%d %H:%M:%S")
        end_time_str = end_time.strftime("%Y-%m-%d %H:%M:%S")
        duration_str = str(work_duration)

        data = {
            'Фамилия': [last_name],
            'Имя': [first_name],
            'Отчество': [patronymic],
            'Время начала': [start_time_str],
            'Время окончания': [end_time_str],
            'Продолжительность работы': [duration_str]
        }
        df = pd.DataFrame(data)

        try:
            existing_data = pd.read_excel("work_data.xlsx")
            df = pd.concat([existing_data, df], ignore_index=True)
        except FileNotFoundError:
            pass

        df.to_excel("work_data.xlsx", index=False)



if __name__ == "__main__":
    root = tk.Tk()
    app = WorkTimerApp(root)
    root.mainloop()
