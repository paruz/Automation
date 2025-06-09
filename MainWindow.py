import customtkinter
import tkinter as tk
from tkinter import filedialog
import os

customtkinter.set_appearance_mode("Light")
customtkinter.set_default_color_theme("green")

class ExcelWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel Processor")
        self.geometry("800x600")
        self.minsize(400, 300)
        self.configure(bg='#f8f8f8')

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_frame = customtkinter.CTkFrame(self, corner_radius=15, fg_color='#f8f8f8')
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure((0, 1, 2), weight=1)
        self.main_frame.grid_rowconfigure((0, 1, 2), weight=1)
        self.main_frame.grid_rowconfigure(3, weight=0)  # Нижняя строка для кнопки Run

        # Upload section grouped in a subframe
        self.upload_frame = customtkinter.CTkFrame(self.main_frame, fg_color='#f8f8f8')
        self.upload_frame.grid(row=0, column=0, rowspan=3, padx=20, pady=20, sticky="nw")

        self.upload_label = customtkinter.CTkLabel(
            self.upload_frame,
            text="Upload Excel File",
            font=customtkinter.CTkFont(size=16, weight="bold"))
        self.upload_label.pack(anchor="w", pady=(0, 2))

        self.upload_btn = customtkinter.CTkButton(
            self.upload_frame,
            text="📂 Upload File",
            command=self.upload_file,
            width=150,
            height=40,
            font=customtkinter.CTkFont(size=14, weight="bold"),
            fg_color="#d4a017",
            hover_color="#a67b0d")
        self.upload_btn.pack(anchor="w", pady=0)

        self.file_name_label = customtkinter.CTkLabel(
            self.upload_frame,
            text="",
            font=customtkinter.CTkFont(size=14, weight="normal"))
        self.file_name_label.pack(anchor="w", pady=(2, 0))

        # Dummy middle column (empty) чтобы Run был по центру между 0 и 2
        self.main_frame.grid_columnconfigure(1, weight=1)

        # Dynamic options frame (правый столбец)
        self.dynamic_frame = customtkinter.CTkFrame(self.main_frame, corner_radius=15, fg_color='#f8f8f8')
        self.dynamic_frame.grid(row=0, column=2, rowspan=3, padx=20, pady=20, sticky="nsew")
        self.dynamic_frame.grid_columnconfigure(0, weight=1)

        self.add_btn = customtkinter.CTkButton(
            self.dynamic_frame,
            text="Add Options",
            command=self.add_dynamic_buttons,
            width=150,
            height=40,
            font=customtkinter.CTkFont(size=14, weight="bold"),
            fg_color="#d4a017",
            hover_color="#a67b0d")
        # Прижать к правому краю фрейма:
        self.add_btn.grid(row=0, column=0, padx=(0, 0), pady=10, sticky="ne")

        self.dynamic_label = customtkinter.CTkLabel(
            self.dynamic_frame,
            text="",
            font=customtkinter.CTkFont(size=14, weight="bold"))

        self.dynamic_btn_frame = customtkinter.CTkFrame(self.dynamic_frame, fg_color='#f8f8f8')
        self.dynamic_btn_frame.grid(row=1, column=0, padx=0, pady=10, sticky="nsew")
        self.dynamic_btn_frame.grid_columnconfigure(0, weight=1)
        self.dynamic_buttons = []

        # Run process button (ниже всех, по центру между колонками 0 и 2)
        self.run_btn = customtkinter.CTkButton(
            self.main_frame,
            text="Run Process",
            command=self.run_process,
            width=200,
            height=50,
            font=customtkinter.CTkFont(size=16, weight="bold"),
            fg_color="#d4a017",
            hover_color="#a67b0d")
        self.run_btn.grid(row=3, column=0, columnspan=3, pady=10, sticky="s")
        # Чтобы кнопка была по центру горизонтально (с шириной 200), используем columnspan=3 и sticky="s" (юг)
        # self.main_frame.grid_columnconfigure(1, weight=1) сделает центрирование

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_name_label.configure(text=os.path.basename(file_path))

    def run_process(self):
        print("Processing started...")

    def add_dynamic_buttons(self):
        self.dynamic_label.grid(row=1, column=0, padx=0, pady=10, sticky="n")
        self.dynamic_btn_frame.grid(row=2, column=0, padx=0, pady=10, sticky="nsew")

        for btn in self.dynamic_buttons:
            btn.destroy()
        self.dynamic_buttons = []

        for i in range(3):
            btn = customtkinter.CTkButton(
                self.dynamic_btn_frame,
                text=f"Option {i+1}",
                command=lambda x=i: self.option_selected(x),
                width=150,
                height=40,
                font=customtkinter.CTkFont(size=14, weight="bold"),
                fg_color="#d4a017",
                hover_color="#a67b0d")
            btn.grid(row=i, column=0, pady=5, padx=0, sticky="ew")
            self.dynamic_buttons.append(btn)

    def option_selected(self, index):
        self.dynamic_label.configure(text=f"Selected Option {index+1}")

if __name__ == "__main__":
    app = ExcelWindow()
    app.mainloop()
