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


class ExcelWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel Processor with Selenium Automation")
        self.geometry("800x600")
        self.minsize(400, 300)
        self.configure(bg='#f8f8f8')

        self.file_path = None
        self.selected_option = "Option 1"

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_frame = customtkinter.CTkFrame(self, corner_radius=15, fg_color='#f8f8f8')
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure((0, 1, 2), weight=1)
        self.main_frame.grid_rowconfigure((0, 1, 2), weight=1)
        self.main_frame.grid_rowconfigure(3, weight=0)

        # Upload section
        self.upload_frame = customtkinter.CTkFrame(self.main_frame, fg_color='#f8f8f8')
        self.upload_frame.grid(row=0, column=0, rowspan=3, padx=20, pady=20, sticky="nw")

        self.upload_label = customtkinter.CTkLabel(self.upload_frame, text="Upload Excel File",
                                                   font=customtkinter.CTkFont(size=16, weight="bold"))
        self.upload_label.pack(anchor="w", pady=(0, 2))

        self.upload_btn = customtkinter.CTkButton(self.upload_frame, text="📂 Upload File", command=self.upload_file,
                                                  width=150, height=40,
                                                  font=customtkinter.CTkFont(size=14, weight="bold"),
                                                  fg_color="#d4a017", hover_color="#a67b0d")
        self.upload_btn.pack(anchor="w", pady=0)

        self.file_name_label = customtkinter.CTkLabel(self.upload_frame, text="",
                                                      font=customtkinter.CTkFont(size=14))
        self.file_name_label.pack(anchor="w", pady=(2, 0))

        # Dynamic frame with dropdown
        self.dynamic_frame = customtkinter.CTkFrame(self.main_frame, corner_radius=15, fg_color='#f8f8f8')
        self.dynamic_frame.grid(row=0, column=2, rowspan=3, padx=20, pady=20, sticky="nsew")
        self.dynamic_frame.grid_columnconfigure(0, weight=1)

        self.option_var = tk.StringVar(value="Option 1")
        self.option_menu = customtkinter.CTkComboBox(
            self.dynamic_frame,
            values=["Option 1", "Option 2", "Option 3"],
            variable=self.option_var,
            width=200,
            font=customtkinter.CTkFont(size=14),
            command=self.option_selected)
        self.option_menu.grid(row=0, column=0, padx=5, pady=10)

        self.selected_label = customtkinter.CTkLabel(self.dynamic_frame, text="Selected: Option 1",
                                                     font=customtkinter.CTkFont(size=14, weight="bold"))
        self.selected_label.grid(row=1, column=0, padx=5, pady=5)

        # Run process button
        self.run_btn = customtkinter.CTkButton(self.main_frame, text="Run Process", command=self.run_process,
                                               width=200, height=50,
                                               font=customtkinter.CTkFont(size=16, weight="bold"),
                                               fg_color="#d4a017", hover_color="#a67b0d")
        self.run_btn.grid(row=3, column=0, columnspan=3, pady=10, sticky="s")

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path = file_path
            self.file_name_label.configure(text=os.path.basename(file_path), text_color="black")

    def option_selected(self, choice=None):
        self.selected_option = self.option_var.get()
        self.selected_label.configure(text=f"Selected: {self.selected_option}")

    def run_process(self):
        if not self.file_path:
            self.file_name_label.configure(text="Please upload an Excel file first!", text_color="red")
            return

        try:
            df = pd.read_excel(self.file_path)
            df.columns = df.columns.str.strip()
        except Exception as e:
            self.file_name_label.configure(text=f"Error reading file: {e}", text_color="red")
            return

        # Закрыть основное окно
        self.destroy()

        # Открыть окно статуса
        status_window = customtkinter.CTk()
        status_window.title("Processing")
        status_window.geometry("500x300")
        status_window.configure(bg='#f8f8f8')

        file_label = customtkinter.CTkLabel(status_window, text=f"File: {os.path.basename(self.file_path)}",
                                            font=customtkinter.CTkFont(size=14, weight="bold"))
        file_label.grid(row=0, column=0, padx=20, pady=10, sticky="w")

        option_label = customtkinter.CTkLabel(status_window, text=f"Option: {self.selected_option}",
                                              font=customtkinter.CTkFont(size=14, weight="bold"))
        option_label.grid(row=0, column=1, padx=20, pady=10, sticky="e")

        progress = customtkinter.CTkProgressBar(status_window)
        progress.grid(row=1, column=0, columnspan=2, padx=20, pady=30, sticky="ew")
        progress.set(0)

        control_frame = customtkinter.CTkFrame(status_window, fg_color="#f8f8f8")
        control_frame.grid(row=2, column=0, columnspan=2, pady=20)

        pause_event = Event()
        stop_event = Event()

        pause_btn = customtkinter.CTkButton(control_frame, text="Pause", command=lambda: pause_event.clear())
        pause_btn.grid(row=0, column=0, padx=10)

        resume_btn = customtkinter.CTkButton(control_frame, text="Resume", command=lambda: pause_event.set())
        resume_btn.grid(row=0, column=1, padx=10)

        stop_btn = customtkinter.CTkButton(control_frame, text="Stop", command=lambda: stop_event.set())
        stop_btn.grid(row=0, column=2, padx=10)

        def automation_task():
            try:
                driver = webdriver.Chrome()
            except Exception as e:
                print(f"Selenium error: {e}")
                return

            pause_event.set()
            stop_event.clear()

            total = len(df)
            for i, row in df.iterrows():
                if stop_event.is_set():
                    break
                pause_event.wait()

                first_name = str(row.get("First Name", "")).strip()
                last_name = str(row.get("Last Name", "")).strip()

                driver.get("https://www.selenium.dev/selenium/web/web-form.html")
                time.sleep(1)

                input_elements = driver.find_elements(By.TAG_NAME, "input")
                if len(input_elements) >= 2:
                    input_elements[0].clear()
                    input_elements[0].send_keys(first_name)
                    input_elements[1].clear()
                    input_elements[1].send_keys(last_name)

                progress.set((i + 1) / total)
                time.sleep(1)

            driver.quit()
            print("Process completed.")

        Thread(target=automation_task, daemon=True).start()
        status_window.mainloop()


if __name__ == "__main__":
    app = ExcelWindow()
    app.mainloop()
