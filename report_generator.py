"""Окно tkinter и генератор учебного отчёта по воронке подбора."""

from __future__ import annotations

import argparse
import queue
import random
import secrets
import threading
import tkinter as tk
import webbrowser
from dataclasses import dataclass
from datetime import date, datetime
from tkinter import messagebox, ttk

from google_sheets import GoogleSheets, SheetsError, load_config


VACANCIES = {
    "Продажи": ["Менеджер по продажам", "Менеджер по работе с клиентами", "Руководитель продаж", "Аккаунт-менеджер"],
    "IT": ["Python-разработчик", "Тестировщик", "Системный аналитик", "Системный администратор"],
    "Офис и поддержка": ["Бухгалтер", "HR-специалист", "Офис-менеджер", "Специалист поддержки"],
}
DEPARTMENTS = ["Все направления", *VACANCIES]


@dataclass(frozen=True)
class ReportOptions:
    start: date
    end: date
    department: str
    count: int

    def validate(self):
        if self.start > self.end:
            raise ValueError("Начало периода должно быть не позже окончания.")
        if self.department not in DEPARTMENTS:
            raise ValueError("Выберите направление из списка.")
        available = sum(len(v) for v in VACANCIES.values()) if self.department == DEPARTMENTS[0] else len(VACANCIES[self.department])
        if not 1 <= self.count <= available:
            raise ValueError(f"Для выбранного направления можно выбрать от 1 до {available} вакансий.")


def generate_data(options: ReportOptions, rng: random.Random | None = None) -> list[list]:
    options.validate()
    rng = rng or random.Random()
    pool = [(role, department) for department, roles in VACANCIES.items()
            if options.department in (DEPARTMENTS[0], department) for role in roles]
    rows = []
    days = (options.end - options.start).days + 1
    for index, (role, department) in enumerate(rng.sample(pool, options.count), start=1):
        plan = rng.randint(1, 4)
        applications = rng.randint(0, min(300, days * 5))
        interviews = rng.randint(0, min(applications, 35))
        offers = rng.randint(0, min(interviews, plan * 2))
        hires = rng.randint(0, min(offers, plan))
        rows.append([index, role, department, plan, applications, interviews, offers, hires])
    return rows


def color(hex_value: str) -> dict:
    value = hex_value.lstrip("#")
    return {component: int(value[index * 2:index * 2 + 2], 16) / 255
            for index, component in enumerate(("red", "green", "blue"))}


def build_report(options: ReportOptions, rows: list[list], *, sheet_id: int, title: str) -> tuple[list[dict], int]:
    """Создаёт запросы данных, формул, размеров, цветов и объединений."""
    options.validate()
    if len(rows) != options.count:
        raise ValueError("Число строк не совпадает с параметрами отчёта.")
    last_data_row = 10 + len(rows)  # Нумерация строк в формулах начинается с 1.
    total_row = last_data_row + 1
    values = [
        [],
        ["Отчёт по подбору сотрудников"],
        [f"Период: {options.start:%d.%m.%Y} — {options.end:%d.%m.%Y}. Направление: {options.department}."],
        ["Учебный отчёт. Все данные сгенерированы случайно."],
        [],
        ["Отклики", "", "Собеседования", "", "Предложения", "", "Выходы на работу"],
        [f"=E{total_row}", "", f"=F{total_row}", "", f"=G{total_row}", "", f"=H{total_row}"],
        [],
        ["Воронка по вакансиям"],
        ["№", "Вакансия", "Направление", "План найма", "Отклики", "Собеседования", "Предложения", "Выходы", "Отклик → выход"],
    ]
    for row_number, row in enumerate(rows, start=11):
        values.append([*row, f'=IF(E{row_number}=0,"",H{row_number}/E{row_number})'])
    values.append(["Итого", "", "", *[f"=SUM({col}11:{col}{last_data_row})" for col in "DEFGH"],
                   f'=IF(E{total_row}=0,"",H{total_row}/E{total_row})'])

    def grid(r1, r2, c1=0, c2=9):
        return {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    def repeat(r1, r2, fmt, c1=0, c2=9):
        return {"repeatCell": {"range": grid(r1, r2, c1, c2),
                               "cell": {"userEnteredFormat": fmt},
                               "fields": ",".join(f"userEnteredFormat.{key}" for key in fmt)}}

    def cell(value):
        if isinstance(value, (int, float)):
            entered = {"numberValue": value}
        elif isinstance(value, str) and value.startswith("="):
            entered = {"formulaValue": value}
        else:
            entered = {"stringValue": str(value)}
        return {"userEnteredValue": entered}

    operations = [
        {"addSheet": {"properties": {"sheetId": sheet_id, "title": title, "index": 0,
            "gridProperties": {"rowCount": max(40, total_row + 3), "columnCount": 9, "hideGridlines": True}}}},
        {"updateCells": {"start": {"sheetId": sheet_id, "rowIndex": 0, "columnIndex": 0},
            "rows": [{"values": [cell(value) for value in row]} for row in values], "fields": "userEnteredValue"}},
        repeat(0, total_row, {"textFormat": {"fontFamily": "Arial", "fontSize": 10, "foregroundColor": color("243247")},
                              "verticalAlignment": "MIDDLE", "wrapStrategy": "CLIP"}),
        repeat(1, 2, {"textFormat": {"fontFamily": "Arial", "fontSize": 16, "bold": True, "foregroundColor": color("243247")}}),
        repeat(2, 4, {"textFormat": {"fontFamily": "Arial", "fontSize": 10, "foregroundColor": color("64748B")}}),
        repeat(5, 7, {"backgroundColor": color("EDF3F9")}),
        repeat(6, 7, {"textFormat": {"fontFamily": "Arial", "fontSize": 16, "bold": True, "foregroundColor": color("243247")}}),
        repeat(8, 9, {"textFormat": {"fontFamily": "Arial", "fontSize": 10, "bold": True}}),
        repeat(9, 10, {"backgroundColor": color("243247"), "horizontalAlignment": "CENTER", "wrapStrategy": "WRAP",
                       "textFormat": {"fontFamily": "Arial", "fontSize": 10, "bold": True, "foregroundColor": color("FFFFFF")}}),
        repeat(10, total_row, {"horizontalAlignment": "RIGHT", "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}, 3, 8),
        repeat(10, total_row, {"horizontalAlignment": "RIGHT", "numberFormat": {"type": "PERCENT", "pattern": "0.0%"}}, 8, 9),
        repeat(total_row - 1, total_row, {"backgroundColor": color("E4EDF6"),
               "textFormat": {"fontFamily": "Arial", "fontSize": 10, "bold": True}}),
        {"setBasicFilter": {"filter": {"range": grid(9, total_row - 1)}}},
        {"updateBorders": {"range": grid(1, 2), "bottom": {"style": "SOLID", "color": color("CBD5E1")}}},
    ]
    for r1, r2, c1, c2 in [(1, 2, 0, 9), (2, 3, 0, 9), (3, 4, 0, 9), (8, 9, 0, 9), (total_row - 1, total_row, 0, 3)]:
        operations.append({"mergeCells": {"range": grid(r1, r2, c1, c2), "mergeType": "MERGE_ALL"}})
    for row in (5, 6):
        for c1, c2 in ((0, 2), (2, 4), (4, 6), (6, 9)):
            operations.append({"mergeCells": {"range": grid(row, row + 1, c1, c2), "mergeType": "MERGE_ALL"}})
    for row in range(10, total_row - 1):
        if row % 2:
            operations.append(repeat(row, row + 1, {"backgroundColor": color("F5F7FA")}))
    for index, width in enumerate((42, 270, 165, 100, 100, 135, 120, 95, 135)):
        operations.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": index, "endIndex": index + 1}, "properties": {"pixelSize": width}, "fields": "pixelSize"}})
    for r1, r2, height in ((0, 1, 15), (1, 2, 40), (2, 4, 26), (5, 6, 28), (6, 7, 36), (8, 9, 26), (9, 10, 44), (10, total_row, 30)):
        operations.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": r1, "endIndex": r2}, "properties": {"pixelSize": height}, "fields": "pixelSize"}})
    return operations, total_row


def write_report(options: ReportOptions, progress=lambda text: None) -> dict:
    rows = generate_data(options)
    config = load_config()
    sheet_id = secrets.randbelow(2_000_000_000) + 1
    title = f"Подбор {datetime.now():%d.%m.%Y %H-%M-%S} {secrets.token_hex(2)}"
    operations, total_row = build_report(options, rows, sheet_id=sheet_id, title=title)
    progress("Записываем данные и оформляем новый лист…")
    with GoogleSheets(**config) as sheets:
        sheets.batch_update(operations)
        progress("Отчёт записан. Проверяем итоговые показатели…")
        try:
            quoted_title = "'" + title.replace("'", "''") + "'"
            actual = sheets.read_range(f"{quoted_title}!D{total_row}:H{total_row}")
            expected = [[sum(row[column] for row in rows) for column in range(3, 8)]]
            verified = actual == expected
        except SheetsError:
            verified = False
    return {"title": title, "sheet_id": sheet_id, "verified": verified,
            "url": f"https://docs.google.com/spreadsheets/d/{config['spreadsheet_id']}/edit#gid={sheet_id}",
            "vacancies": len(rows), "hires": sum(row[7] for row in rows)}


class ReportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Отчёт по подбору сотрудников")
        self.geometry("780x740")
        self.minsize(740, 720)
        self.configure(bg="#F4F6FA")
        self.result_url = ""
        self.events = queue.Queue()
        self.busy = False
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#F4F6FA")
        style.configure("TLabel", background="#F4F6FA", foreground="#243247", font=("Segoe UI", 11))
        style.configure("Title.TLabel", font=("Segoe UI", 19, "bold"))
        style.configure("Note.TLabel", foreground="#64748B", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 11), padding=(14, 10))
        style.configure("Primary.TButton", background="#315C9B", foreground="white")
        style.map("Primary.TButton", background=[("active", "#254A80"), ("disabled", "#94A3B8")])
        content = ttk.Frame(self, padding=28)
        content.pack(fill="both", expand=True)
        content.columnconfigure(1, weight=1)
        ttk.Label(content, text="Отчёт по подбору сотрудников", style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(content, text="Выберите период и сформируйте отчёт в Google Таблице.", style="Note.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(7, 24))
        today = date.today()
        self.start_var = tk.StringVar(value=today.replace(day=1).strftime("%d.%m.%Y"))
        self.end_var = tk.StringVar(value=today.strftime("%d.%m.%Y"))
        self.department_var = tk.StringVar(value=DEPARTMENTS[0])
        self.count_var = tk.StringVar(value="6")
        self.inputs = []
        for row, (label, variable) in enumerate((("Начало периода", self.start_var), ("Конец периода", self.end_var)), start=2):
            ttk.Label(content, text=label).grid(row=row, column=0, sticky="w", pady=9)
            entry = ttk.Entry(content, textvariable=variable, font=("Segoe UI", 11))
            entry.grid(row=row, column=1, sticky="ew", padx=(22, 0), pady=9)
            self.inputs.append(entry)
        ttk.Label(content, text="Формат даты: ДД.ММ.ГГГГ", style="Note.TLabel").grid(row=4, column=1, sticky="w", padx=(22, 0))
        ttk.Label(content, text="Направление").grid(row=5, column=0, sticky="w", pady=12)
        self.department = ttk.Combobox(content, textvariable=self.department_var, values=DEPARTMENTS, state="readonly", font=("Segoe UI", 11))
        self.department.grid(row=5, column=1, sticky="ew", padx=(22, 0), pady=12)
        self.department.bind("<<ComboboxSelected>>", self.adjust_count)
        ttk.Label(content, text="Количество вакансий").grid(row=6, column=0, sticky="w", pady=9)
        self.count = ttk.Spinbox(content, from_=1, to=12, textvariable=self.count_var, font=("Segoe UI", 11))
        self.count.grid(row=6, column=1, sticky="ew", padx=(22, 0), pady=9)
        ttk.Label(content, text="Учебный симулятор: все показатели генерируются случайно.\nКаждый отчёт сохраняется на новом листе таблицы.", style="Note.TLabel").grid(row=7, column=0, columnspan=2, sticky="w", pady=(18, 15))
        self.generate_button = ttk.Button(content, text="Сформировать отчёт", style="Primary.TButton", command=self.start_report)
        self.generate_button.grid(row=8, column=0, columnspan=2, sticky="ew")
        self.progress = ttk.Progressbar(content, mode="indeterminate")
        self.progress.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.status_var = tk.StringVar(value="Готово к созданию отчёта.")
        ttk.Label(content, textvariable=self.status_var, wraplength=660).grid(row=10, column=0, columnspan=2, sticky="w", pady=12)
        self.open_button = ttk.Button(content, text="Открыть готовый отчёт", command=self.open_report, state="disabled")
        self.open_button.grid(row=11, column=0, columnspan=2, sticky="ew")
        self.after(100, self.poll_events)
        self.protocol("WM_DELETE_WINDOW", self.close_window)

    def adjust_count(self, _=None):
        maximum = 12 if self.department_var.get() == DEPARTMENTS[0] else 4
        self.count.configure(to=maximum)
        try:
            current = int(self.count_var.get())
        except ValueError:
            current = maximum
        self.count_var.set(str(max(1, min(current, maximum))))

    def read_options(self):
        try:
            start = datetime.strptime(self.start_var.get().strip(), "%d.%m.%Y").date()
            end = datetime.strptime(self.end_var.get().strip(), "%d.%m.%Y").date()
        except ValueError:
            raise ValueError("Введите существующие даты в формате ДД.ММ.ГГГГ, например 01.10.2026.") from None
        try:
            count = int(self.count_var.get())
        except ValueError:
            raise ValueError("Количество вакансий должно быть целым числом.") from None
        options = ReportOptions(start, end, self.department_var.get(), count)
        options.validate()
        return options

    def set_busy(self, busy):
        self.busy = busy
        for entry in self.inputs:
            entry.configure(state="disabled" if busy else "normal")
        self.department.configure(state="disabled" if busy else "readonly")
        self.count.configure(state="disabled" if busy else "normal")
        self.generate_button.configure(state="disabled" if busy else "normal")
        if busy:
            self.open_button.configure(state="disabled")
            self.progress.start(12)
        else:
            self.progress.stop()
            self.open_button.configure(state="normal" if self.result_url else "disabled")

    def start_report(self):
        if self.busy:
            return
        try:
            options = self.read_options()
        except ValueError as error:
            messagebox.showwarning("Проверьте поля", str(error), parent=self)
            return
        self.result_url = ""
        self.set_busy(True)
        self.status_var.set("Подключаемся к Google Таблице…")
        threading.Thread(target=self.worker, args=(options,), daemon=True).start()

    def worker(self, options):
        try:
            result = write_report(options, lambda message: self.events.put(("progress", message)))
            self.events.put(("success", result))
        except (SheetsError, ValueError) as error:
            self.events.put(("error", str(error)))
        except Exception:
            self.events.put(("error", "Не удалось завершить операцию. Проверьте таблицу перед повторным запуском."))

    def poll_events(self):
        while True:
            try:
                event, value = self.events.get_nowait()
            except queue.Empty:
                break
            if event == "progress":
                self.status_var.set(value)
            elif event == "success":
                self.result_url = value["url"]
                self.set_busy(False)
                if value["verified"]:
                    self.status_var.set(f"Отчёт создан и проверен. Вакансий: {value['vacancies']}. Выходов на работу: {value['hires']}.")
                else:
                    self.status_var.set("Отчёт записан. Итоговые показатели нужно проверить в таблице.")
            else:
                self.set_busy(False)
                self.status_var.set("Создание отчёта не подтверждено. Проверьте таблицу перед повтором.")
                messagebox.showerror("Не удалось подтвердить результат", value, parent=self)
        self.after(100, self.poll_events)

    def open_report(self):
        if self.result_url:
            webbrowser.open(self.result_url)

    def close_window(self):
        if self.busy and not messagebox.askyesno("Идёт запись отчёта", "Запись ещё выполняется. Закрыть окно? Перед повтором проверьте таблицу.", parent=self):
            return
        self.destroy()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--generate", action="store_true", help="Создать отчёт без окна")
    parser.add_argument("--start", help="Начало периода: ГГГГ-ММ-ДД")
    parser.add_argument("--end", help="Конец периода: ГГГГ-ММ-ДД")
    parser.add_argument("--department", default=DEPARTMENTS[0])
    parser.add_argument("--count", type=int, default=6)
    args = parser.parse_args()
    if args.generate:
        today = date.today()
        try:
            options = ReportOptions(date.fromisoformat(args.start) if args.start else today.replace(day=1),
                                    date.fromisoformat(args.end) if args.end else today, args.department, args.count)
            result = write_report(options, print)
        except (SheetsError, ValueError) as error:
            parser.exit(1, str(error) + "\n")
        print("Отчёт:", result["url"])
        print("Итоги проверены:", result["verified"])
    else:
        ReportApp().mainloop()


if __name__ == "__main__":
    main()
