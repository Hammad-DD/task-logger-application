import pandas as pd
from datetime import datetime
import time
import customtkinter as tk
from tkinter import messagebox


class TaskTrackerApp:
    def __init__(self):
        self.task_data = None

        self.root = tk.CTk()
        self.root.title("Task Tracker")
        self.root.geometry('300x200')
        try:
            tk.set_default_color_theme("stupid.json")

        except FileNotFoundError:
            pass

        self.task_name_label = tk.CTkLabel(self.root, text="Enter the task name:")
        self.task_name_label.pack()

        self.task_name_entry = tk.CTkEntry(self.root)
        self.task_name_entry.pack(pady=3)  # remember to add padding
        self.task_name_entry.bind("<Return>", self.enter_key)

        self.start_button = tk.CTkButton(self.root, text="Start Task", command=self.on_start_button_click)
        self.start_button.pack()

        self.end_button = tk.CTkButton(self.root, text="End Task", command=self.on_end_button_click, state=tk.DISABLED)
        self.end_button.pack()

        self.total_time_label = tk.CTkLabel(self.root, text="Total Time Spent: 0 Hrs")
        self.total_time_label.pack()

        self.task_names_label = tk.CTkLabel(self.root, text="Tasks:")
        self.task_names_label.pack()
        self.error = False

    def start_task(self):
        task_name = self.task_name_entry.get()
        start_time = time.time()
        start_date = datetime.now()
        day_of_week = start_date.strftime('%A')
        self.task_data = (task_name, start_date, start_time, day_of_week)
        self.start_button.configure(state=tk.DISABLED)
        self.task_name_entry.configure(state=tk.DISABLED)
        self.end_button.configure(state=tk.NORMAL)
        self.total_time_label.configure(
            text=f'Currently working on:\n {task_name}\n Started at: {start_date.strftime("%H:%M:%S")}')

    def end_task(self):
        if self.task_data:
            task_name, start_date, start_time, day_of_week = self.task_data
            end_time = time.time()
            end_date = datetime.now()
            time_spent = (((end_time - start_time) / 60) / 60) # time in hours
            time_spent = round(time_spent, 3)
            self.log_to_excel(task_name, start_date, end_date, day_of_week, start_time, end_time, time_spent)
            if (self.error == False):
                messagebox.showinfo("Task Complete", "Task logged successfully!")
                self.start_button.configure(state=tk.NORMAL)
                self.end_button.configure(state=tk.DISABLED)
                self.task_name_entry.configure(state=tk.NORMAL)
                self.update_total_time()
                self.update_task_names()

    def log_to_excel(self, task_name, start_date, end_date, day_of_week, start_time, end_time, time_spent):
        data = {
            'Task': [task_name],
            'Starting DateTime': [start_date],
            'Day': [day_of_week],
            'End DateTime': [end_date],
            'Start Time': [start_time],
            'End Time': [end_time],
            'Time Spent(H)': [time_spent]
        }

        try:
            df = pd.read_excel('task_log.xlsx')  # If the file already exists
        except FileNotFoundError:
            df = pd.DataFrame(columns=['Task', 'Starting DateTime', 'End DateTime', 'Day', 'Start Time', 'End Time',
                                       'Time Spent(H)'])

        new_entry = pd.DataFrame(data)
        df = pd.concat([df, new_entry], ignore_index=True)
        try:
            df.to_excel('task_log.xlsx', index=False, engine='openpyxl')
            self.error = False
        except PermissionError:
            self.error = True
            messagebox.showinfo("Task NOT LOGGED",
                                "Please make sure the file is not open elsewhere ")

    def update_total_time(self):
        df = pd.read_excel('task_log.xlsx')
        total_time = df['Time Spent(H)'].sum()
        hours = int(total_time)
        minutes = round((total_time % 1) * 60)
        self.total_time_label.configure(text=f"Total Time Spent: {hours} Hrs {minutes} m")

    def update_task_names(self):
        df = pd.read_excel('task_log.xlsx')
        task_names = df['Task'].unique()
        task_names_str = "\n".join(task_names)
        self.task_names_label.configure(text=f"Tasks:\n{task_names_str}")

    def on_start_button_click(self):
        self.start_task()

    def enter_key(self, event):
        self.on_start_button_click()

    def on_end_button_click(self):
        self.end_task()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = TaskTrackerApp()
    app.run()
