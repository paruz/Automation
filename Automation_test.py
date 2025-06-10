import customtkinter
import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from threading import Thread, Event

customtkinter.set_appearance_mode("Light")
customtkinter.set_default_color_theme("green")


class ExcelApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel Processor with Selenium Automation")
        self.geometry("800x600")
        self.minsize(400, 300)

        self.file_path = None
        self.selected_option = "Option 1"

        self.pause_event = Event()
        self.stop_event = Event()

        # Сюда будут записываться dict с ключами: "First Name", "Last Name", "Status"
        self.report_data = []

        self.init_main_screen()

    def init_main_screen(self):
        self.main_frame = customtkinter.CTkFrame(self, corner_radius=15)
        self.main_frame.pack(expand=True, fill="both", padx=20, pady=20)

        self.upload_label = customtkinter.CTkLabel(
            self.main_frame, text="Upload Excel File",
            font=customtkinter.CTkFont(size=16, weight="bold")
        )
        self.upload_label.pack(pady=(10, 2))

        self.upload_btn = customtkinter.CTkButton(
            self.main_frame, text="📂 Upload File", command=self.upload_file,
            width=150, height=40,
            font=customtkinter.CTkFont(size=14, weight="bold"),
            fg_color="#d4a017", hover_color="#a67b0d"
        )
        self.upload_btn.pack()

        self.file_name_label = customtkinter.CTkLabel(
            self.main_frame, text="",
            font=customtkinter.CTkFont(size=14)
        )
        self.file_name_label.pack(pady=(2, 10))

        self.option_var = tk.StringVar(value="Option 1")
        self.option_menu = customtkinter.CTkComboBox(
            self.main_frame,
            values=["Option 1", "Option 2", "Option 3"],
            variable=self.option_var,
            width=200,
            font=customtkinter.CTkFont(size=14),
            command=self.option_selected
        )
        self.option_menu.pack(pady=10)

        self.selected_label = customtkinter.CTkLabel(
            self.main_frame, text="Selected: Option 1",
            font=customtkinter.CTkFont(size=14, weight="bold")
        )
        self.selected_label.pack()

        self.run_btn = customtkinter.CTkButton(
            self.main_frame, text="Start Automation", command=self.start_processing,
            width=200, height=50,
            font=customtkinter.CTkFont(size=16, weight="bold"),
            fg_color="#d4a017", hover_color="#a67b0d"
        )
        self.run_btn.pack(pady=20)

        # Элементы для отображения прогресса и управления
        self.progress = customtkinter.CTkProgressBar(self.main_frame)
        self.time_label = customtkinter.CTkLabel(
            self.main_frame, text="Estimated time left: --",
            font=customtkinter.CTkFont(size=14)
        )
        self.controls = customtkinter.CTkFrame(self.main_frame, fg_color="#f8f8f8")

        self.pause_btn = customtkinter.CTkButton(
            self.controls, text="Pause",
            command=lambda: self.pause_event.clear()
        )
        self.continue_btn = customtkinter.CTkButton(
            self.controls, text="Continue",
            command=lambda: self.pause_event.set()
        )
        self.stop_btn = customtkinter.CTkButton(
            self.controls, text="Stop", command=lambda: self.stop_event.set()
        )

        self.pause_btn.pack(side="left", padx=10)
        self.continue_btn.pack(side="left", padx=10)
        self.stop_btn.pack(side="left", padx=10)

    def upload_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file_path = path
            self.file_name_label.configure(text=os.path.basename(path), text_color="black")

    def option_selected(self, choice=None):
        self.selected_option = self.option_var.get()
        self.selected_label.configure(text=f"Selected: {self.selected_option}")

    def start_processing(self):
        if not self.file_path:
            self.file_name_label.configure(text="Please upload an Excel file first!", text_color="red")
            return

        try:
            self.df = pd.read_excel(self.file_path)
            self.df.columns = self.df.columns.str.strip()
        except Exception as e:
            self.file_name_label.configure(text=f"Error reading file: {e}", text_color="red")
            return

        self.pause_event.set()
        self.stop_event.clear()
        self.report_data = []  # Сброс отчёта при каждом запуске

        # Скрываем исходные элементы
        self.upload_label.pack_forget()
        self.upload_btn.pack_forget()
        self.file_name_label.pack_forget()
        self.option_menu.pack_forget()
        self.selected_label.pack_forget()
        self.run_btn.pack_forget()

        # Показываем элементы прогресса и кнопок управления
        self.progress.pack(pady=(10, 5), padx=40, fill="x")
        self.progress.set(0)
        self.time_label.pack(pady=(0, 20))
        self.controls.pack(pady=20)

        Thread(target=self.automation_task, daemon=True).start()

    def format_time(self, seconds):
        minutes = int(seconds) // 60
        seconds = int(seconds) % 60
        return f"{minutes} min {seconds} sec" if minutes > 0 else f"{seconds} sec"

    def automation_task(self):
        try:
            driver = webdriver.Chrome()
        except Exception as e:
            self.time_label.configure(text=f"WebDriver error: {e}")
            return

        total = len(self.df)
        start_time = time.time()

        for i, row in self.df.iterrows():
            # Извлекаем только необходимые данные: имя и фамилия
            first_name_raw = row.get("First Name", "")
            last_name_raw = row.get("Last Name", "")
            first_name = "" if pd.isna(first_name_raw) else str(first_name_raw).strip()
            last_name = "" if pd.isna(last_name_raw) else str(last_name_raw).strip()

            # Формируем базовую структуру строки отчёта
            report_row = {"First Name": first_name, "Last Name": last_name}

            if self.stop_event.is_set():
                report_row["Status"] = "Processing stopped by user"
                self.report_data.append(report_row)
                break

            self.pause_event.wait()

            try:
                if not first_name or not last_name:
                    raise ValueError("Missing first name or last name")
                driver.get("https://www.selenium.dev/selenium/web/web-form.html")
                time.sleep(0.5)

                inputs = driver.find_elements(By.TAG_NAME, "input")
                if len(inputs) < 2:
                    raise Exception("Not enough input fields found on the page")

                inputs[0].clear()
                inputs[0].send_keys(first_name)
                inputs[1].clear()
                inputs[1].send_keys(last_name)
            except Exception as e:
                report_row["Status"] = f"Error: {str(e)}"
            else:
                report_row["Status"] = "Success"

            self.report_data.append(report_row)

            self.progress.set((i + 1) / total)
            elapsed = time.time() - start_time
            avg_time = elapsed / (i + 1)
            remaining_time = avg_time * (total - (i + 1))
            self.time_label.configure(text=f"Estimated time left: {self.format_time(remaining_time)}")
            time.sleep(0.5)

        driver.quit()
        if self.stop_event.is_set():
            self.time_label.configure(text="Processing stopped by user")
        else:
            self.time_label.configure(text="Completed successfully ✅")

        self.progress.pack_forget()
        self.controls.pack_forget()

        # Добавляем кнопку для скачивания отчёта
        self.download_btn = customtkinter.CTkButton(
            self.main_frame, text="Download Report", command=self.download_report,
            width=200, height=50,
            font=customtkinter.CTkFont(size=16, weight="bold"),
            fg_color="#d4a017", hover_color="#a67b0d"
        )
        self.download_btn.pack(pady=20)

        # Добавляем под кнопкой Download Report кнопку для возвращения к начальному полю
        self.return_btn = customtkinter.CTkButton(
            self.main_frame, text="Return to Main", command=self.reset_ui,
            width=200, height=50,
            font=customtkinter.CTkFont(size=16, weight="bold"),
            fg_color="#d4a017", hover_color="#a67b0d"
        )
        self.return_btn.pack(pady=10)

    def download_report(self):
        file_save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            title="Save Report",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_save_path:
            try:
                df_report = pd.DataFrame(self.report_data, columns=["First Name", "Last Name", "Status"])
                df_report.to_excel(file_save_path, index=False)
                self.time_label.configure(text=f"Report saved: {os.path.basename(file_save_path)}")
            except Exception as e:
                self.time_label.configure(text=f"Error saving report: {e}")

    def reset_ui(self):
        # Удаляем кнопки отчёта
        if hasattr(self, 'download_btn'):
            self.download_btn.pack_forget()
        if hasattr(self, 'return_btn'):
            self.return_btn.pack_forget()
        # Скрываем метку с информацией по времени и очищаем её текст
        self.time_label.pack_forget()
        self.time_label.configure(text="")

        # Возвращаем начальные элементы
        self.upload_label.pack(pady=(10, 2))
        self.upload_btn.pack()
        self.file_name_label.configure(text="")
        self.file_name_label.pack(pady=(2, 10))
        self.option_menu.pack(pady=10)
        self.selected_label.pack()
        self.run_btn.pack(pady=20)

        # Сбрасываем путь файла
        self.file_path = None


if __name__ == "__main__":
    app = ExcelApp()
    app.mainloop()
