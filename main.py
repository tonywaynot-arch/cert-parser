import os
import re
import pdfplumber
from tkinter import Tk, filedialog

# Русские месяцы
MONTHS = {
    "января": "01",
    "февраля": "02",
    "марта": "03",
    "апреля": "04",
    "мая": "05",
    "июня": "06",
    "июля": "07",
    "августа": "08",
    "сентября": "09",
    "октября": "10",
    "ноября": "11",
    "декабря": "12",
}

def normalize_date(text):
    # 01.02.2024
    m = re.search(r"\b(\d{2}\.\d{2}\.\d{4})\b", text)
    if m:
        return m.group(1)

    # 1 января 2024 г.
    m = re.search(r"(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})", text, re.IGNORECASE)
    if m:
        day = m.group(1).zfill(2)
        month = MONTHS[m.group(2).lower()]
        year = m.group(3)
        return f"{day}.{month}.{year}"

    return "дата не найдена"

def extract_text_from_pdf(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def extract_material(text):
    m = re.search(r"Продукция[:\s\-]*([^\n]+)", text)
    if m:
        return m.group(1).strip()
    return "Материал не найден"

def extract_certificate_number(text):
    patterns = [
        r"(EAES\.[A-Z0-9\.\-]+)",
        r"(ROSS\.[A-Z0-9\.\-]+)",
        r"(RU\s*C-RU\.[A-Z0-9\.\-]+)",
        r"(SSBK\.[A-Z0-9\.\-]+)",
    ]
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            return m.group(1).replace(" ", "")
    return "номер не найден"

def main():
    Tk().withdraw()
    folder = filedialog.askdirectory(title="Выберите папку с PDF файлами")

    if not folder:
        return

    output_lines = []

    for file in os.listdir(folder):
        if file.lower().endswith(".pdf"):
            path = os.path.join(folder, file)
            text = extract_text_from_pdf(path)

            material = extract_material(text)
            cert_number = extract_certificate_number(text)
            date = normalize_date(text)

            line = f". {material} (Сертификат №{cert_number}, от {date});"
            output_lines.append(line)

    result_path = os.path.join(folder, "result.txt")
    with open(result_path, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))

    print("Готово. Файл result.txt создан.")

if __name__ == "__main__":
    main()
