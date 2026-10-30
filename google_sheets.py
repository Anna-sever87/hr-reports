"""Переиспользуемый класс для Google Sheets API: CRUD и оформление.

Доступ предоставляется отдельному файлу таблицы. Google Drive API не нужен.
Обычный текст передаётся с RAW и не превращается в формулы.
"""

from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import quote

import requests
from google.auth.exceptions import GoogleAuthError
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials


class SheetsError(Exception):
    """Понятное сообщение пользователю без содержимого ключа."""


class GoogleSheets:
    BASE_URL = "https://sheets.googleapis.com/v4/spreadsheets"
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    def __init__(self, credentials_path: str | Path, spreadsheet_id: str):
        self.spreadsheet_id = spreadsheet_id
        if not spreadsheet_id or not all(c.isalnum() or c in "-_" for c in spreadsheet_id):
            raise SheetsError("В настройках нужен ID таблицы, а не полная ссылка.")
        key_path = Path(credentials_path)
        if not key_path.is_file():
            raise SheetsError("Файл ключа не найден. Проверьте путь в настройках.")
        try:
            credentials = Credentials.from_service_account_file(str(key_path), scopes=self.SCOPES)
            self.session = AuthorizedSession(credentials)
        except (ValueError, OSError, GoogleAuthError):
            raise SheetsError("Не удалось прочитать ключ сервисного аккаунта.") from None

    def close(self):
        self.session.close()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()

    def _request(self, method: str, suffix: str = "", **kwargs) -> dict:
        url = f"{self.BASE_URL}/{self.spreadsheet_id}{suffix}"
        try:
            response = self.session.request(method, url, timeout=(10, 55), **kwargs)
        except GoogleAuthError:
            raise SheetsError("Google не принял ключ. Проверьте, что ключ активен и часы компьютера верны.") from None
        except requests.RequestException:
            raise SheetsError(
                "Нет ответа от Google. Проверьте интернет. Если отправляли отчёт, "
                "сначала проверьте таблицу: запрос мог выполниться."
            ) from None
        if not response.ok:
            messages = {
                400: "Google отклонил параметры запроса. Проверьте диапазон и данные.",
                401: "Google не принял учётные данные. Проверьте ключ.",
                403: "Доступ запрещён. Проверьте включение Sheets API и доступ Editor для сервисного аккаунта.",
                404: "Таблица или лист не найдены. Проверьте ID таблицы и предоставленный доступ.",
                429: "Слишком много запросов. Подождите минуту и повторите попытку.",
            }
            message = messages.get(response.status_code, "Ошибка сервиса Google. Попробуйте позже.")
            raise SheetsError(f"{message} Код: {response.status_code}.")
        return response.json()

    def metadata(self) -> dict:
        return self._request("GET", params={"fields": "spreadsheetId,properties(title),sheets(properties)"})

    def get_sheet_names(self) -> list[str]:
        return [sheet["properties"]["title"] for sheet in self.metadata().get("sheets", [])]

    def read_range(self, cell_range: str, *, formulas: bool = False) -> list[list]:
        result = self._request(
            "GET", f"/values/{quote(cell_range, safe='')}",
            params={"valueRenderOption": "FORMULA" if formulas else "UNFORMATTED_VALUE"},
        )
        return result.get("values", [])

    def read_all_cells(self, sheet_name: str | None = None) -> list[list]:
        if sheet_name is None:
            names = self.get_sheet_names()
            if not names:
                return []
            sheet_name = names[0]
        return self.read_range("'" + sheet_name.replace("'", "''") + "'")

    def update_range(self, cell_range: str, values: list[list]) -> dict:
        return self._request(
            "PUT", f"/values/{quote(cell_range, safe='')}",
            params={"valueInputOption": "RAW"}, json={"values": values},
        )

    def update_cell(self, cell: str, value) -> dict:
        return self.update_range(cell, [[value]])

    def append_rows(self, cell_range: str, values: list[list]) -> dict:
        return self._request(
            "POST", f"/values/{quote(cell_range, safe='')}:append",
            params={"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"},
            json={"values": values},
        )

    def clear_range(self, cell_range: str) -> dict:
        """Удаляет значения в диапазоне, сохраняя оформление."""
        return self._request("POST", f"/values/{quote(cell_range, safe='')}:clear", json={})

    def batch_update(self, operations: list[dict]) -> dict:
        """Весь пакет применяется атомарно: данные и оформление вместе."""
        return self._request("POST", ":batchUpdate", json={"requests": operations})

    def create_sheet(self, title: str) -> int:
        result = self.batch_update([{"addSheet": {"properties": {"title": title}}}])
        return result["replies"][0]["addSheet"]["properties"]["sheetId"]

    def delete_sheet(self, sheet_id: int) -> dict:
        return self.batch_update([{"deleteSheet": {"sheetId": sheet_id}}])


def load_config(config_path: str | Path | None = None) -> dict:
    base = Path(__file__).resolve().parent
    path = Path(config_path) if config_path else base / "config.local.json"
    try:
        config = json.loads(path.read_text(encoding="utf-8-sig"))
        credentials_path = Path(config["credentials_path"])
        if not credentials_path.is_absolute():
            credentials_path = path.resolve().parent / credentials_path
        return {"credentials_path": credentials_path, "spreadsheet_id": config["spreadsheet_id"]}
    except (OSError, ValueError, KeyError, TypeError):
        raise SheetsError("Не удалось прочитать config.local.json. Скопируйте config.example.json и заполните настройки.") from None


if __name__ == "__main__":
    try:
        with GoogleSheets(**load_config()) as sheets:
            print("Листы:", ", ".join(sheets.get_sheet_names()))
            for index, row in enumerate(sheets.read_all_cells(), start=1):
                print(f"Строка {index}: {row}")
    except SheetsError as error:
        print(error)
        raise SystemExit(1)
