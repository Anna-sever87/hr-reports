"""Собрать архив для сдачи по явному списку файлов, без ключа и настроек."""

from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

BASE = Path(__file__).resolve().parent
FILES = [
    "google_sheets.py", "report_generator.py", "requirements.txt", "README.md",
    "config.example.json", ".gitignore", "start.cmd", "package_homework.py", "verify_project.py",
    ".vscode/settings.json", ".vscode/launch.json",
]


def main():
    output = BASE.parent / "ДЗ 2 урок HR Reports.zip"
    for filename in FILES:
        if not (BASE / filename).is_file():
            raise FileNotFoundError(filename)
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for filename in FILES:
            archive.write(BASE / filename, f"HR Reports/{filename}")
    print("Создан архив:", output)
    print("Ключ, личные настройки и виртуальное окружение не включены.")


if __name__ == "__main__":
    main()
