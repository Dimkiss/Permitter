import sys
from pathlib import Path
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from src.excel_parser import InvoiceParser
from src.docx_builder import DocxBuilder
from src.email_sender import OutlookMailer


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("RPA: разрешение на перемещение МЦ")
        self.root.geometry("1200x750")
        self.root.minsize(1000, 650)

        self.parser = InvoiceParser()
        self.builder = DocxBuilder()
        self.mailer = OutlookMailer()
        self.items = []

        if getattr(sys, "frozen", False):
            self.base_dir = Path(sys._MEIPASS)
        else:
            self.base_dir = Path(__file__).resolve().parent.parent

        self.template_path = self.base_dir / "template.docx"

        self.invoice_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.cwd()))
        self.email_to = tk.StringVar(value="security@example.com")
        self.subject = tk.StringVar(value="Разрешение на перемещение материальных ценностей")

        self.operation_type = tk.StringVar(value="Ввоз/вывоз")
        self.date_value = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))
        self.time_from = tk.StringVar(value="09:00")
        self.time_to = tk.StringVar(value="18:00")
        self.location = tk.StringVar(
            value="664040, Иркутская область, г. Иркутск, ул. Розы Люксембург, д. 184"
        )
        self.responsible_name = tk.StringVar(value="Иванов И.И.")
        self.phone = tk.StringVar(value="+7 900 000-00-00")

        self.preview_before_send = tk.BooleanVar(value=False)

        self.build_ui()

        self.root.after(100, self.set_initial_pane_ratio)

    def log(self, text):
        self.log_box.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')}  {text}\n")
        self.log_box.see(tk.END)
        self.root.update_idletasks()

    def set_initial_pane_ratio(self):
        self.root.update_idletasks()
        total_width = self.paned.winfo_width()
        if total_width > 0:
            self.paned.sashpos(0, int(total_width * 0.6))

    def build_ui(self):
        container = ttk.Frame(self.root, padding=10)
        container.pack(fill="both", expand=True)

        self.paned = ttk.PanedWindow(container, orient="horizontal")
        self.paned.pack(fill="both", expand=True)

        left = ttk.Frame(self.paned)
        right = ttk.Frame(self.paned)

        self.paned.add(left, weight=3)
        self.paned.add(right, weight=2)

        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        files_frame = ttk.LabelFrame(left, text="Файлы", padding=10)
        files_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(files_frame, text="Счёт Excel").grid(row=0, column=0, sticky="w")
        ttk.Entry(files_frame, textvariable=self.invoice_path).grid(
            row=1, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Button(files_frame, text="Выбрать", command=self.pick_invoice).grid(
            row=1, column=1
        )

        ttk.Label(files_frame, text="Папка для сохранения").grid(
            row=2, column=0, sticky="w", pady=(10, 0)
        )
        ttk.Entry(files_frame, textvariable=self.output_dir).grid(
            row=3, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Button(files_frame, text="Выбрать", command=self.pick_output_dir).grid(
            row=3, column=1
        )

        files_frame.columnconfigure(0, weight=1)

        form_frame = ttk.LabelFrame(left, text="Данные", padding=10)
        form_frame.grid(row=1, column=0, sticky="new")

        fields = [
            ("Тип операции", self.operation_type),
            ("Дата", self.date_value),
            ("Время с", self.time_from),
            ("Время до", self.time_to),
            ("Локация", self.location),
            ("Ответственный", self.responsible_name),
            ("Телефон", self.phone),
            ("Email", self.email_to),
            ("Тема", self.subject),
        ]

        for i, (label, var) in enumerate(fields):
            ttk.Label(form_frame, text=label).grid(
                row=i, column=0, sticky="w", pady=3, padx=(0, 10)
            )

            if var is self.operation_type:
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=var,
                    values=["Ввоз/вывоз", "Внос/вынос", "Перемещение"],
                    state="readonly",
                )
            else:
                widget = ttk.Entry(form_frame, textvariable=var)

            widget.grid(row=i, column=1, sticky="ew", pady=3)

        form_frame.columnconfigure(1, weight=1)

        bottom_left = ttk.Frame(left)
        bottom_left.grid(row=2, column=0, sticky="sew", pady=(10, 0))
        bottom_left.columnconfigure(0, weight=1)

        controls_frame = ttk.Frame(bottom_left)
        controls_frame.grid(row=0, column=0, sticky="ew")

        ttk.Checkbutton(
            controls_frame,
            text="Открывать перед отправкой",
            variable=self.preview_before_send,
        ).pack(anchor="w", pady=(0, 8))

        ttk.Button(
            controls_frame,
            text="Проверить Outlook",
            command=self.check_outlook,
        ).pack(fill="x", pady=3)

        ttk.Button(
            controls_frame,
            text="Прочитать Excel",
            command=self.parse_invoice,
        ).pack(fill="x", pady=3)

        ttk.Button(
            controls_frame,
            text="Создать Word",
            command=self.build_docx,
        ).pack(fill="x", pady=3)

        ttk.Button(
            controls_frame,
            text="Отправить Email",
            command=self.send_email,
        ).pack(fill="x", pady=3)

        preview_frame = ttk.LabelFrame(right, text="Найденные позиции", padding=10)
        preview_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)

        columns = ("name", "quantity")
        self.tree = ttk.Treeview(preview_frame, columns=columns, show="headings")
        self.tree.heading("name", text="Наименование")
        self.tree.heading("quantity", text="Количество")
        self.tree.column("name", width=260, anchor="w")
        self.tree.column("quantity", width=110, anchor="center", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")

        tree_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.grid(row=0, column=1, sticky="ns")

        log_frame = ttk.LabelFrame(right, text="Лог", padding=10)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_box = tk.Text(log_frame, wrap="word")
        self.log_box.grid(row=0, column=0, sticky="nsew")

        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=log_scroll.set)
        log_scroll.grid(row=0, column=1, sticky="ns")

    def pick_invoice(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xls *.xlsx")])
        if path:
            self.invoice_path.set(path)

    def pick_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def check_outlook(self):
        ok, msg = self.mailer.check()
        self.log(msg)
        if ok:
            messagebox.showinfo("Outlook", msg)
        else:
            messagebox.showwarning("Outlook", msg)

    def parse_invoice(self):
        path = self.invoice_path.get().strip()
        if not path:
            messagebox.showwarning("Ошибка", "Выбери Excel-файл.")
            return

        try:
            self.log("Чтение Excel...")
            self.items = self.parser.parse_items(path)

            for row in self.tree.get_children():
                self.tree.delete(row)

            for item in self.items:
                self.tree.insert("", tk.END, values=(item["name"], item["quantity"]))

            self.log(f"Найдено позиций: {len(self.items)}")
        except Exception as e:
            self.log(f"Ошибка разбора Excel: {e}")
            messagebox.showerror("Ошибка", str(e))

    def collect_doc_data(self):
        return {
            "operation_type": self.operation_type.get().strip(),
            "date": self.date_value.get().strip(),
            "time_from": self.time_from.get().strip(),
            "time_to": self.time_to.get().strip(),
            "location": self.location.get().strip(),
            "responsible_name": self.responsible_name.get().strip(),
            "phone": self.phone.get().strip(),
        }

    def get_output_docx_path(self):
        safe_date = datetime.now().strftime("%Y%m%d_%H%M%S")
        return str(Path(self.output_dir.get()) / f"permit_{safe_date}.docx")

    def build_docx(self):
        if not self.items:
            messagebox.showwarning("Ошибка", "Сначала прочитай Excel.")
            return

        if not self.template_path.exists():
            messagebox.showerror("Ошибка", "Файл template.docx не найден в корне проекта.")
            return

        try:
            output_path = self.get_output_docx_path()
            self.log("Формирование Word...")

            self.builder.build(
                template_path=self.template_path,
                output_path=output_path,
                data=self.collect_doc_data(),
                items=self.items,
            )

            self.log(f"Документ сохранён: {output_path}")
            messagebox.showinfo("Готово", f"Документ создан:\n{output_path}")

        except Exception as e:
            self.log(f"Ошибка формирования Word: {e}")
            self.log(traceback.format_exc())
            messagebox.showerror("Ошибка", str(e))

    def send_email(self):
        if not self.items:
            messagebox.showwarning("Ошибка", "Сначала прочитай Excel.")
            return

        if not self.template_path.exists():
            messagebox.showerror("Ошибка", "Файл template.docx не найден.")
            return

        try:
            output_path = self.get_output_docx_path()

            self.log("Создание Word перед отправкой...")
            self.builder.build(
                template_path=self.template_path,
                output_path=output_path,
                data=self.collect_doc_data(),
                items=self.items,
            )

            ok, msg = self.mailer.check()
            self.log(msg)

            if not ok:
                messagebox.showwarning("Outlook", msg)
                return

            body = (
                f"Добрый день.\n\n"
                f"Разрешение на {self.operation_type.get().lower()} материальных ценностей.\n"
                f"Дата: {self.date_value.get()}\n"
                f"Время: {self.time_from.get()} - {self.time_to.get()}\n"
                f"Локация: {self.location.get()}\n"
                f"Ответственный: {self.responsible_name.get()}\n"
                f"Телефон: {self.phone.get()}\n\n"
                f"Документ во вложении."
            )

            result = self.mailer.send(
                to_email=self.email_to.get(),
                subject=self.subject.get(),
                body=body,
                attachment_path=output_path,
                display_only=self.preview_before_send.get(),
            )

            self.log(result)
            messagebox.showinfo("Готово", result)

        except Exception as e:
            self.log(f"Ошибка отправки: {e}")
            self.log(traceback.format_exc())
            messagebox.showerror("Ошибка", str(e))