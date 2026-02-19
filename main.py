import os
import re
import tkinter as tk
from tkinter import filedialog
from docx import Document
from datetime import datetime

def format_date(date_text):
    months = {
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

    match = re.search(r"(\d{1,2})\s+([а-я]+)\s+(\d{4})", date_text.lower())
    if match:
        day, month, year = match.groups()
        month_num = months.get(month, "01")
        return f"{int(day):02d}.{month_num}.{year}"
    return ""

def extract_data(text):
    product = ""
    cert_number = ""
    date = ""

    product_match = re.search(r"Продукция[:\s]*(.*)", text)
    if product_match:
        product = product_match.group(1).strip()

    cert_match = re.search(
        r"(Сертификат соответствия|паспорт|Документ о качестве)[^\n]*?(EAES\S+|ROSS\S+|RU C-RU\S+|SSBK\S+)",
        text,
        re.IGNORECASE,
    )
    if cert_match:
        cert_number = cert_match.group(2)

    date_match = re.search(r"\d{1,2}\s+[а-я]+\s+\d{4}", text.lower())
    if date_match:
        date = format_date(date_match.group())

    return product, cert_number, date

def process_folder(folder_path):
    results = []

    for file in os.listdir(folder_path):
        if file.endswith(".docx"):
            doc = Document(os.path.join(folder_path, file))
            full_text = "\n".join([p.text for p in doc.paragraphs])
            product, cert_number, date = extract_data(full_text)

            if product:
                line = f"{product} (Сертификат №{cert_number}, от {date});"
                results.append(line)

    return results

def main():
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Выберите папку с файлами")

    if not folder_path:
        return

    results = process_folder(folder_path)

    output_file = os.path.join(folder_path, "result.txt")
    with open(output_file, "w", encoding="utf-8") as f:
        for line in results:
            f.write(line + "\n")

    print("Готово! Файл result.txt создан.")

if __name__ == "__main__":
    main()
