import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import subprocess
import threading
import json
import os
import re
import sys
from pathlib import Path

class OrionParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OrionMAXParser GUI")
        
        # Set window size
        self.root.geometry("900x700")
        
        self.config_path = Path("config.json")
        self.selected_json = tk.StringVar()
        self.month_var = tk.StringVar()
        self.weather_markers_var = tk.StringVar()
        self.tech_markers_var = tk.StringVar()
        
        self.create_widgets()
        self.load_initial_config()

    def create_widgets(self):
        # Header
        header_label = tk.Label(self.root, text="Управление OrionMAXParser", font=("Arial", 16, "bold"), pady=10)
        header_label.pack()

        # Settings frame
        file_frame = tk.LabelFrame(self.root, text="Настройки входа", padx=10, pady=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(file_frame, text="JSON файл:").grid(row=0, column=0, sticky="w")
        self.entry_json = tk.Entry(file_frame, textvariable=self.selected_json, width=70)
        self.entry_json.grid(row=0, column=1, padx=5)
        tk.Button(file_frame, text="Обзор", command=self.browse_file).grid(row=0, column=2)
        
        tk.Label(file_frame, text="Месяц (ГГГГ-ММ):").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(file_frame, textvariable=self.month_var, width=20).grid(row=1, column=1, sticky="w", padx=5)
        tk.Label(file_frame, text="(извлекается автоматически из имени файла)", fg="gray").grid(row=1, column=1, sticky="e", padx=5)

        # Filters frame
        filter_frame = tk.LabelFrame(self.root, text="Фильтры (через запятую)", padx=10, pady=10)
        filter_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(filter_frame, text="Тех. поломки:").grid(row=0, column=0, sticky="w")
        tk.Entry(filter_frame, textvariable=self.tech_markers_var, width=85).grid(row=0, column=1, padx=5)
        
        tk.Label(filter_frame, text="Погода/Игнор:").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(filter_frame, textvariable=self.weather_markers_var, width=85).grid(row=1, column=1, padx=5)
        
        btn_save = tk.Button(filter_frame, text="Сохранить настройки", command=self.manual_save, bg="#FF9800", fg="white")
        btn_save.grid(row=2, column=1, sticky="e", pady=5)
        
        # Action frame
        action_frame = tk.Frame(self.root, pady=10)
        action_frame.pack(fill="x", padx=10)
        
        self.btn_parse = tk.Button(action_frame, text="1. Запустить парсинг событий", 
                                   command=self.run_parsing, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), height=2, width=35)
        self.btn_parse.pack(side="left", padx=10)
        
        self.btn_report = tk.Button(action_frame, text="2. Сгенерировать отчет Excel", 
                                    command=self.run_reporting, bg="#2196F3", fg="white", font=("Arial", 10, "bold"), height=2, width=35)
        self.btn_report.pack(side="left", padx=10)
        
        # Console output
        log_header_frame = tk.Frame(self.root)
        log_header_frame.pack(fill="x", padx=10)
        
        tk.Label(log_header_frame, text="Логи выполнения:").pack(side="left")
        tk.Button(log_header_frame, text="Копировать всё", command=self.copy_all_logs, font=("Arial", 8)).pack(side="right")
        
        self.log_area = scrolledtext.ScrolledText(self.root, height=15, font=("Consolas", 9))
        self.log_area.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_area.bind("<Button-3>", self.show_context_menu) # Right click
        
        # Context menu
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Копировать", command=self.copy_selected)
        self.context_menu.add_command(label="Выделить всё", command=self.select_all)
        
        # Status bar
        self.status_var = tk.StringVar(value="Готов")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief="sunken", anchor="w")
        status_bar.pack(side="bottom", fill="x")

    def show_context_menu(self, event):
        self.context_menu.post(event.x_root, event.y_root)

    def copy_selected(self):
        try:
            selected_text = self.log_area.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
        except tk.TclError:
            pass # No selection

    def select_all(self):
        self.log_area.tag_add(tk.SEL, "1.0", tk.END)
        self.log_area.mark_set(tk.INSERT, "1.0")
        self.log_area.see(tk.INSERT)
        return 'break'

    def copy_all_logs(self):
        logs = self.log_area.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(logs)
        self.status_var.set("Логи скопированы в буфер обмена")

    def load_initial_config(self):
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    json_path = cfg.get("json_path", "")
                    self.selected_json.set(json_path)
                    self.extract_month_from_filename(json_path)
                    
                    self.weather_markers_var.set(", ".join(cfg.get("weather_markers", [])))
                    self.tech_markers_var.set(", ".join(cfg.get("technical_markers", [])))
            except Exception as e:
                self.log(f"Ошибка загрузки config.json: {e}")

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            initialdir=".",
            title="Выберите JSON файл экспорта чата",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if file_path:
            # Save relative path if possible
            try:
                rel_path = os.path.relpath(file_path, os.getcwd())
                self.selected_json.set(rel_path)
            except ValueError:
                self.selected_json.set(file_path)
            
            self.extract_month_from_filename(self.selected_json.get())

    def extract_month_from_filename(self, filename):
        # Format: MM.YYYY.json or similar
        match = re.search(r"(\d{2})\.(\d{4})", filename)
        if match:
            mm, yyyy = match.groups()
            self.month_var.set(f"{yyyy}-{mm}")
        else:
            # Try YYYY-MM
            match = re.search(r"(\d{4})-(\d{2})", filename)
            if match:
                yyyy, mm = match.groups()
                self.month_var.set(f"{yyyy}-{mm}")

    def log(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.config(state="disabled")
        self.log_area.see(tk.END)

    def update_config(self):
        if not self.config_path.exists():
            messagebox.showerror("Ошибка", "Файл config.json не найден!")
            return False
        
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            
            cfg["json_path"] = self.selected_json.get()
            
            # Update markers
            weather = [s.strip() for s in self.weather_markers_var.get().split(",") if s.strip()]
            tech = [s.strip() for s in self.tech_markers_var.get().split(",") if s.strip()]
            cfg["weather_markers"] = weather
            cfg["technical_markers"] = tech
            
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            self.log(f"Конфигурация обновлена: json_path = {cfg['json_path']}")
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить config.json: {e}")
            return False

    def manual_save(self):
        if self.update_config():
            messagebox.showinfo("Успех", "Настройки успешно сохранены в config.json")

    def run_command(self, cmd_list, success_msg):
        def target():
            self.btn_parse.config(state="disabled")
            self.btn_report.config(state="disabled")
            self.status_var.set("Выполнение...")
            
            try:
                # Use CREATE_NO_WINDOW if on Windows to avoid flashing cmd
                creation_flags = 0
                if sys.platform == "win32":
                    creation_flags = subprocess.CREATE_NO_WINDOW

                cmd_display = " ".join(f'"{c}"' if " " in str(c) else str(c) for c in cmd_list)
                self.log(f"> {cmd_display}")
                self.log("... ожидание вывода (может занять время для тяжелых файлов) ...")

                # Force unbuffered output so we see logs in real-time
                env = os.environ.copy()
                env["PYTHONUNBUFFERED"] = "1"

                process = subprocess.Popen(
                    cmd_list,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    env=env,
                    bufsize=1, # Line buffered
                    creationflags=creation_flags
                )
                
                # Use a loop that checks for process survival
                while True:
                    line = process.stdout.readline()
                    if not line and process.poll() is not None:
                        break
                    if line:
                        self.log(line.strip())
                
                process.wait()
                if process.returncode == 0:
                    self.log(f"\n✅ {success_msg}")
                    self.status_var.set("Завершено успешно")
                    messagebox.showinfo("Успех", success_msg)
                else:
                    self.log(f"\n❌ Ошибка при выполнении. Код: {process.returncode}")
                    self.status_var.set("Ошибка выполнения")
                    messagebox.showerror("Ошибка", f"Скрипт завершился с ошибкой (код {process.returncode})")
                    
            except Exception as e:
                self.log(f"\n❌ Критическая ошибка: {e}")
                self.status_var.set("Ошибка")
                messagebox.showerror("Критическая ошибка", str(e))
            
            self.btn_parse.config(state="normal")
            self.btn_report.config(state="normal")

        threading.Thread(target=target, daemon=True).start()

    def run_parsing(self):
        if not self.update_config():
            return
        
        month = self.month_var.get()
        if not re.match(r"\d{4}-\d{2}", month):
            messagebox.showwarning("Внимание", "Введите месяц в формате ГГГГ-ММ")
            return
        
        self.log_area.config(state="normal")
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state="disabled")
        self.log(f"--- Запуск парсинга для {month} ---")
        self.log(f"Файл: {self.selected_json.get()}")
        
        # Use current python executable (the one from venv if started via bat)
        python_exe = sys.executable
        cmd = [python_exe, "instructor_daily_attraction_events.py", "--month", month, "--reuse"]
        self.run_command(cmd, "Парсинг событий завершен!")

    def run_reporting(self):
        month = self.month_var.get()
        if not re.match(r"\d{4}-\d{2}", month):
            messagebox.showwarning("Внимание", "Введите месяц в формате ГГГГ-ММ")
            return
        
        self.log(f"\n--- Генерация отчета за {month} ---")
        
        python_exe = sys.executable
        cmd = [python_exe, "generate_report.py", "--month", month]
        self.run_command(cmd, "Отчет Excel сгенерирован!")

if __name__ == "__main__":
    root = tk.Tk()
    # Try to set icon if exists (optional)
    # if os.path.exists("icon.ico"): root.iconbitmap("icon.ico")
    
    app = OrionParserGUI(root)
    root.mainloop()
