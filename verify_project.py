"""Проверка модели данных и CRUD на отдельном временном листе.

Запуск без --live не обращается к Google. --live создаёт один учебный отчёт.
"""

import argparse
import json
import random
import secrets
from datetime import date, datetime
from pathlib import Path

from google_sheets import GoogleSheets, load_config
from report_generator import DEPARTMENTS, ReportOptions, generate_data, write_report


def check_model():
    for department in DEPARTMENTS:
        count = 6 if department == DEPARTMENTS[0] else 4
        options = ReportOptions(date(2026, 10, 1), date(2026, 10, 30), department, count)
        for seed in range(20):
            rows = generate_data(options, random.Random(seed))
            assert len(rows) == count
            assert len({row[1] for row in rows}) == count
            for _, _, _, plan, applications, interviews, offers, hires in rows:
                assert applications >= interviews >= offers >= hires >= 0
                assert 0 <= hires <= plan
    for options in (
        ReportOptions(date(2026, 10, 30), date(2026, 10, 1), DEPARTMENTS[0], 6),
        ReportOptions(date(2026, 10, 1), date(2026, 10, 30), "IT", 6),
        ReportOptions(date(2026, 10, 1), date(2026, 10, 30), DEPARTMENTS[0], 0),
    ):
        try:
            generate_data(options)
        except ValueError:
            pass
        else:
            raise AssertionError("Некорректные параметры должны отклоняться.")
    print("Модель данных и проверка полей: OK")


def check_live():
    with GoogleSheets(**load_config()) as sheets:
        metadata = sheets.metadata()
        print("Подключение к таблице:", metadata["properties"]["title"])
        title = "Проверка CRUD " + secrets.token_hex(4)
        temporary_id = sheets.create_sheet(title)
        quoted = "'" + title + "'"
        try:
            sheets.update_range(f"{quoted}!A1:B2", [["Проверка", "Значение"], ["Отклики", 10]])
            assert sheets.read_range(f"{quoted}!A1:B2") == [["Проверка", "Значение"], ["Отклики", 10]]
            sheets.update_cell(f"{quoted}!B2", 15)
            assert sheets.read_range(f"{quoted}!B2") == [[15]]
            sheets.append_rows(f"{quoted}!A1:B2", [["Собеседования", 4]])
            assert sheets.read_range(f"{quoted}!A3:B3") == [["Собеседования", 4]]
            sheets.update_cell(f"{quoted}!D1", "=1+1")
            assert sheets.read_range(f"{quoted}!D1") == [["=1+1"]], "Текст должен оставаться текстом."
            sheets.clear_range(f"{quoted}!D1")
            assert sheets.read_range(f"{quoted}!D1") == []
        finally:
            sheets.delete_sheet(temporary_id)
        assert title not in sheets.get_sheet_names()
        print("Создание, чтение, добавление, изменение, очистка и удаление временного листа: OK")

    result = write_report(ReportOptions(date(2026, 10, 1), date(2026, 10, 30), DEPARTMENTS[0], 6), print)
    assert result["verified"], "Итоги в Google Таблице должны совпадать с исходными числами."
    with GoogleSheets(**load_config()) as sheets:
        quoted = "'" + result["title"] + "'"
        values = sheets.read_range(f"{quoted}!A1:I20")
        errors = {"#ERROR!", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!"}
        assert not any(isinstance(cell, str) and cell in errors for row in values for cell in row)
        assert sheets.read_range(f"{quoted}!I11", formulas=True) == [['=IF(E11=0,"",H11/E11)']]
        properties = next(sheet["properties"] for sheet in sheets.metadata()["sheets"]
                          if sheet["properties"]["sheetId"] == result["sheet_id"])
        assert properties["gridProperties"]["hideGridlines"]
    print("Формулы, записанные итоги и новый лист: OK")
    print("Готовый отчёт:", result["url"])
    checks = Path(__file__).resolve().parent / ".checks"
    checks.mkdir(exist_ok=True)
    (checks / "verification.json").write_text(json.dumps({
        "checked_at": datetime.now().astimezone().isoformat(),
        "crud": "passed", "model": "passed", "report_formulas": "passed", "report": result,
    }, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--live", action="store_true")
    args = parser.parse_args()
    check_model()
    if args.live:
        check_live()
