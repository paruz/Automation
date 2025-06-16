import customtkinter
import tkinter as tk
from tkinter import filedialog
import os
import time
from threading import Thread, Event
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

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

        self.report_data = []

        self.init_main_screen()

    def init_main_screen(self):
        # Set unified background and smaller window size
        self.configure(fg_color="#f4f4f4")
        self.geometry("600x450")

        self.main_frame = customtkinter.CTkFrame(self, corner_radius=10, fg_color="#f4f4f4")
        self.main_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # Centralized container for all elements
        self.container = customtkinter.CTkFrame(self.main_frame, fg_color="#f4f4f4")
        self.container.pack(expand=True)

        # Upload Excel File
        self.upload_label = customtkinter.CTkLabel(
            self.container, text="Upload Excel File",
            font=customtkinter.CTkFont(size=16, weight="bold"),
            text_color="#222222"
        )
        self.upload_label.pack(pady=(10, 4))

        self.upload_btn = customtkinter.CTkButton(
            self.container, text="📂 Browse File", command=self.upload_file,
            width=220, height=48,
            font=customtkinter.CTkFont(size=15),
            fg_color="#3b82f6", hover_color="#2563eb"
        )
        self.upload_btn.pack(pady=2)

        self.file_name_label = customtkinter.CTkLabel(
            self.container, text="", font=customtkinter.CTkFont(size=13),
            text_color="#555555"
        )
        self.file_name_label.pack(pady=(4, 12))

        # Options ComboBox
        self.option_var = tk.StringVar(value="Option 1")
        self.option_menu = customtkinter.CTkComboBox(
            self.container,
            values=["Option 1", "Option 2", "Option 3"],
            variable=self.option_var,
            width=220,
            font=customtkinter.CTkFont(size=14),
            command=self.option_selected
        )
        self.option_menu.pack(pady=(0, 6))

        self.selected_label = customtkinter.CTkLabel(
            self.container, text="Selected: Option 1",
            font=customtkinter.CTkFont(size=13),
            text_color="#444444"
        )
        self.selected_label.pack(pady=(0, 16))

        # Start Button
        self.run_btn = customtkinter.CTkButton(
            self.container, text="🚀 Start Automation", command=self.start_processing,
            width=250, height=52,
            font=customtkinter.CTkFont(size=16, weight="bold"),
            fg_color="#10b981", hover_color="#059669"
        )
        self.run_btn.pack()

        # Progress and controls (hidden for now but initialized)
        self.progress = customtkinter.CTkProgressBar(self.container, height=16)
        self.progress.set(0)

        self.time_label = customtkinter.CTkLabel(
            self.container, text="Estimated time left: --",
            font=customtkinter.CTkFont(size=13),
            text_color="#333333"
        )

        self.controls = customtkinter.CTkFrame(self.container, fg_color="#f4f4f4")

        self.pause_btn = customtkinter.CTkButton(
            self.controls, text="Pause", command=lambda: self.pause_event.clear(),
            fg_color="#facc15", hover_color="#eab308"
        )
        self.continue_btn = customtkinter.CTkButton(
            self.controls, text="Continue", command=lambda: self.pause_event.set(),
            fg_color="#4ade80", hover_color="#22c55e"
        )
        self.stop_btn = customtkinter.CTkButton(
            self.controls, text="Stop", command=lambda: self.stop_event.set(),
            fg_color="#f87171", hover_color="#ef4444"
        )

        self.pause_btn.pack(side="left", padx=8)
        self.continue_btn.pack(side="left", padx=8)
        self.stop_btn.pack(side="left", padx=8)

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
            self.wb = load_workbook(self.file_path)
            self.ws = self.wb.active
            # Check that there are First Name and Last Name columns in the first row
            headers = [cell.value.strip() if isinstance(cell.value, str) else "" for cell in self.ws[1]]
            if "First Name" not in headers or "Last Name" not in headers:
                raise ValueError("Excel file must contain 'First Name' and 'Last Name' columns")
            self.first_name_col = headers.index("First Name") + 1  # openpyxl is 1-based
            self.last_name_col = headers.index("Last Name") + 1
            self.total_rows = self.ws.max_row - 1  # excluding header
        except Exception as e:
            self.file_name_label.configure(text=f"Error reading file: {e}", text_color="red")
            return

        self.pause_event.set()
        self.stop_event.clear()
        self.report_data = []

        # Hide initial widgets
        self.upload_label.pack_forget()
        self.upload_btn.pack_forget()
        self.file_name_label.pack_forget()
        self.option_menu.pack_forget()
        self.selected_label.pack_forget()
        self.run_btn.pack_forget()

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
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service)
        except Exception as e:
            self.time_label.configure(text=f"WebDriver error: {e}")
            return

        start_time = time.time()

        for i, row in enumerate(self.ws.iter_rows(min_row=2, max_row=self.ws.max_row, values_only=True)):
            first_name_raw = row[self.first_name_col - 1]
            last_name_raw = row[self.last_name_col - 1]

            first_name = "" if first_name_raw is None else str(first_name_raw).strip()
            last_name = "" if last_name_raw is None else str(last_name_raw).strip()

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

            self.progress.set((i + 1) / self.total_rows)
            elapsed = time.time() - start_time
            avg_time = elapsed / (i + 1)
            remaining_time = avg_time * (self.total_rows - (i + 1))
            self.time_label.configure(text=f"Estimated time left: {self.format_time(remaining_time)}")
            time.sleep(0.5)

        driver.quit()

        # Remove old widgets
        self.progress.pack_forget()
        self.controls.pack_forget()
        self.time_label.pack_forget()

        # Finish — third screen
        finish_message = (
            "✅ Completed successfully" if not self.stop_event.is_set()
            else "⛔ Processing stopped by user"
        )

        self.finish_label = customtkinter.CTkLabel(
            self.container, text=finish_message,
            font=customtkinter.CTkFont(size=16, weight="bold"),
            text_color="#10b981" if not self.stop_event.is_set() else "#dc2626"
        )
        self.finish_label.pack(pady=(20, 12))

        self.download_btn = customtkinter.CTkButton(
            self.container, text="💾 Download Report", command=self.download_report,
            width=240, height=48,
            font=customtkinter.CTkFont(size=15, weight="bold"),
            fg_color="#3b82f6", hover_color="#2563eb"
        )
        self.download_btn.pack(pady=(0, 10))

        self.return_btn = customtkinter.CTkButton(
            self.container, text="🔁 Return to Main", command=self.reset_ui,
            width=240, height=48,
            font=customtkinter.CTkFont(size=15, weight="bold"),
            fg_color="#a855f7", hover_color="#9333ea"
        )
        self.return_btn.pack()

    def download_report(self):
        file_save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            title="Save Report",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_save_path:
            try:
                from openpyxl import Workbook
                wb_report = Workbook()
                ws_report = wb_report.active
                ws_report.append(["First Name", "Last Name", "Status"])
                for row in self.report_data:
                    ws_report.append([row.get("First Name", ""), row.get("Last Name", ""), row.get("Status", "")])
                wb_report.save(file_save_path)
                self.time_label.configure(text=f"Report saved: {os.path.basename(file_save_path)}")
            except Exception as e:
                self.time_label.configure(text=f"Error saving report: {e}")

    def reset_ui(self):
        # Remove finish elements
        if hasattr(self, 'download_btn'):
            self.download_btn.pack_forget()
        if hasattr(self, 'return_btn'):
            self.return_btn.pack_forget()
        if hasattr(self, 'finish_label'):
            self.finish_label.pack_forget()

        self.time_label.configure(text="")
        self.file_name_label.configure(text="")

        # Show main screen again
        self.upload_label.pack(pady=(10, 4))
        self.upload_btn.pack(pady=2)
        self.file_name_label.pack(pady=(4, 12))
        self.option_menu.pack(pady=(0, 6))
        self.selected_label.pack(pady=(0, 16))
        self.run_btn.pack()

        # Reset file path
        self.file_path = None


if __name__ == "__main__":
    app = ExcelApp()
    app.mainloop()
