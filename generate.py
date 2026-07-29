#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор пакетов документов Word из реестра Excel.

Запуск: python generate.py
Нужен только Python 3.7+ (входит в Windows 10/11).
"""

import os
import sys
import shutil
import zipfile
import re
import traceback
import xml.etree.ElementTree as ET

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

REGISTRY_FILE = "Реестр_клиентов.xlsx"
TEMPLATE_DIR = "Шаблоны"
OUTPUT_DIR = "Выход"

TEMPLATE_TYPES = ["Договор", "Заявление", "Прилож3", "Счёт", "Лист"]

# Namespace для SpreadsheetML
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def ns(tag):
    return f"{{{SHEET_NS}}}{tag}"


class XlsxReader:
    """Читает xlsx через zipfile + ElementTree (minimally)."""

    def __init__(self, filepath):
        self.z = zipfile.ZipFile(filepath, "r")
        # Прочитаем shared strings один раз
        self.shared_strings = []
        if "xl/sharedStrings.xml" in self.z.namelist():
            try:
                ss_root = ET.fromstring(self.z.read("xl/sharedStrings.xml"))
                for si in ss_root.iter(ns("si")):
                    texts = []
                    for t in si.iter(ns("t")):
                        if t.text:
                            texts.append(t.text)
                    self.shared_strings.append("".join(texts))
            except Exception:
                pass

        # Определим соответствие sheet ID -> имя файла
        wb = ET.fromstring(self.z.read("xl/workbook.xml"))
        self.sheets = []  # [(sheet_name, rel_id)]
        for sheet in wb.iter(ns("sheet")):
            self.sheets.append((sheet.get("name"), sheet.get("sheetId")))

        # rels
        rels = ET.fromstring(self.z.read("xl/_rels/workbook.xml.rels"))
        self.id_to_file = {}
        for rel in rels:
            rid = rel.get("Id")
            target = rel.get("Target")
            if target and rid:
                self.id_to_file[rid] = target

    def get_sheet_by_index(self, idx):
        """Возвращает (sheet_name, [[cell_value,...],...])"""
        if idx >= len(self.sheets):
            return None, []
        name, sid = self.sheets[idx]
        # Найдём файл через rels (идём по порядку — первый rel, второй и т.д.
        # Проще: файлы листов называются sheet1.xml, sheet2.xml обычно
        sheet_file = f"xl/worksheets/sheet{idx + 1}.xml"
        if sheet_file not in self.z.namelist():
            # fallback: ищем по rels
            wb_rels = ET.fromstring(self.z.read("xl/_rels/workbook.xml.rels"))
            for rel in wb_rels:
                target = rel.get("Target")
                rid = rel.get("Id")
                # Найдём соответствие sheetId -> rid (они идут в том же порядке обычно)
                # Но для надёжности просто возьмём по индексу
            return name, []

        ws_xml = self.z.read(sheet_file)
        root = ET.fromstring(ws_xml)
        rows = []
        for row_el in root.iter(ns("row")):
            r = int(row_el.get("r", 0))
            # Расширяем список rows до нужного размера
            while len(rows) < r:
                rows.append([])
            cells = row_el.findall(ns("c"))
            # Определяем макс колонку по атрибуту r
            max_col = 0
            row_values = []
            for c in cells:
                ref = c.get("r", "")
                col_idx = self._col_to_num(ref)
                if col_idx > max_col:
                    max_col = col_idx
            row_values = [""] * max_col
            for c in cells:
                ref = c.get("r", "")
                col_idx = self._col_to_num(ref) - 1  # 0-based
                val = self._cell_value(c)
                row_values[col_idx] = val
            rows[r - 1] = row_values
        return name, rows

    def _col_to_num(self, ref):
        """A1 -> 1, B2 -> 2, AA1 -> 27"""
        col_str = ""
        for ch in ref:
            if ch.isalpha():
                col_str += ch
            else:
                break
        num = 0
        for ch in col_str:
            num = num * 26 + (ord(ch.upper()) - ord("A") + 1)
        return num

    def _cell_value(self, c_el):
        t_attr = c_el.get("t", "")
        v_el = c_el.find(ns("v"))
        is_el = c_el.find(ns("is"))  # inlineStr
        if is_el is not None:
            texts = []
            for t in is_el.iter(ns("t")):
                if t.text:
                    texts.append(t.text)
            return "".join(texts)
        if v_el is None:
            return ""
        val = v_el.text or ""
        if t_attr == "s":
            # shared string
            try:
                idx = int(val)
                if 0 <= idx < len(self.shared_strings):
                    return self.shared_strings[idx]
            except ValueError:
                pass
            return ""
        return val

    def close(self):
        self.z.close()


# ── DOCX PROCESSING ──────────────────────────────────────────────────

def resolve_placeholder(ph_content: str, mapping: dict) -> str:
    """
    Находит лучшее значение для плейсхолдера по содержимому между { }.
    Может искать внутри содержимого подстроку-key и наоборот.
    """
    ph_normalized = ph_content.strip().lower()
    best_value = None
    best_score = 0

    for key, val in mapping.items():
        key_norm = key.strip().lower()
        # Точное совпадение
        if ph_normalized == key_norm:
            return val
        # Подстрока — key содержится в ph
        if key_norm in ph_normalized:
            if len(key_norm) > best_score:
                best_score = len(key_norm)
                best_value = val
        # Подстрока — ph содержится в key
        elif ph_normalized in key_norm:
            if len(ph_normalized) > best_score:
                best_score = len(ph_normalized)
                best_value = val
    return best_value


def replace_placeholders_in_xml(xml: str, mapping: dict) -> str:
    """
    Заменяет ВСЕ плейсхолдеры вида {anything} в XML docx.
    Работает даже если placeholder разбит на несколько <w:r>.
    """
    wt_re = re.compile(r'(<w:t(?:\s+xml:space="preserve")?>)([^<]*)(</w:t>)')

    for _ in range(2000):
        wts = list(wt_re.finditer(xml))
        if not wts:
            break
        flat = "".join(m.group(2) for m in wts)

        # Ищем первый {placeholder}
        match = re.search(r'\{([^{}]+)\}', flat)
        if not match:
            break

        ph_full = match.group(0)   # {дата договора «__» августа 2026 г.}
        ph_content = match.group(1)
        replacement_value = resolve_placeholder(ph_content, mapping)
        if replacement_value is None:
            replacement_value = ""  # placeholder не распознан — удаляем

        start_char = match.start()
        end_char = match.end()

        # Map char positions → wt indices
        char_pos = 0
        start_wt_idx = None
        end_wt_idx = None
        for i, m in enumerate(wts):
            text_len = len(m.group(2))
            if start_wt_idx is None and char_pos <= start_char < char_pos + text_len:
                start_wt_idx = i
            if start_wt_idx is not None and char_pos < end_char <= char_pos + text_len:
                end_wt_idx = i
                break
            char_pos += text_len

        if start_wt_idx is None or end_wt_idx is None:
            # Не смогли привязать — пропускаем этот placeholder
            # Чтобы не зациклиться, заменим его на пустое на flat-уровне (не трогаем XML)
            # Но мы не можем, потому что XML не изменился... просто выходим
            break

        start_xml_pos = wts[start_wt_idx].start()
        search_start = max(0, start_xml_pos - 5000)
        r_open = xml.rfind("<w:r", search_start, start_xml_pos)
        if r_open < 0:
            r_open = start_xml_pos

        end_xml_pos = wts[end_wt_idx].end()
        r_close = xml.find("</w:r>", end_xml_pos)
        if r_close >= 0:
            r_close += len("</w:r>")
        else:
            r_close = end_xml_pos

        # Расширяем влево до <w:proofErr> если прямо перед
        pre_start = max(0, r_open - 300)
        last_proof = xml.rfind("<w:proofErr", pre_start, r_open)
        if last_proof >= 0:
            proof_end = xml.find(">", last_proof) + 1
            if proof_end <= r_open:
                r_open = last_proof

        safe_val = (
            (replacement_value or "")
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )
        replacement = f'<w:r><w:t xml:space="preserve">{safe_val}</w:t></w:r>'
        xml = xml[:r_open] + replacement + xml[r_close:]

    return xml


def process_docx(template_path: str, output_path: str, row_data: dict):
    """Копирует docx и заменяет плейсхолдеры."""
    shutil.copy(template_path, output_path)
    zin = zipfile.ZipFile(template_path, "r")
    zout = zipfile.ZipFile(output_path, "w")

    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            xml = data.decode("utf-8")
            xml = replace_placeholders_in_xml(xml, row_data)
            data = xml.encode("utf-8")
        zout.writestr(item, data)

    zin.close()
    zout.close()


def abbreviate(fio: str) -> str:
    if not fio:
        return ""
    parts = fio.replace(".", " ").split()
    parts = [p for p in parts if p]
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    elif len(parts) == 2:
        return f"{parts[0]} {parts[1][0]}."
    elif len(parts) == 1:
        return parts[0]
    return fio


def main():
    if not os.path.isfile(REGISTRY_FILE):
        print(f'ОШИБКА: "{REGISTRY_FILE}" не найден в папке:\n  {os.path.abspath(SCRIPT_DIR)}')
        input("Нажмите Enter для выхода...")
        sys.exit(1)

    reader = XlsxReader(REGISTRY_FILE)
    total_rows = 0
    total_files = 0
    errors = []

    sheet_count = len(reader.sheets)
    for sheet_idx in range(sheet_count):
        sheet_name, rows = reader.get_sheet_by_index(sheet_idx)
        if not rows or len(rows) < 2:
            print(f"Пропуск листа '{sheet_name}' (нет данных)")
            continue

        set_number = "1"
        m = re.search(r"(\d+)", sheet_name or "")
        if m:
            set_number = m.group(1)

        print(f"\nЛист: {sheet_name} (Набор {set_number})")

        headers = rows[0]
        for row_idx in range(1, len(rows)):
            row_values = rows[row_idx]
            if not row_values:
                continue

            # Собираем row_data
            row_data = {}
            contract_number = ""
            for col_idx, header in enumerate(headers):
                if not header:
                    continue
                val = row_values[col_idx] if col_idx < len(row_values) else ""
                row_data[header] = str(val) if val is not None else ""
                if header.lower() == "номер договора":
                    contract_number = str(val) if val else ""

            if not contract_number:
                print(f"  Строка {row_idx + 1}: пропуск (нет номера договора)")
                continue

            total_rows += 1
            out_folder = os.path.join(OUTPUT_DIR, contract_number)
            os.makedirs(out_folder, exist_ok=True)

            # Производные поля
            fio1 = row_data.get("ФИО_подписанта_1", "")
            fio2 = row_data.get("ФИО_подписанта_2", "")
            row_data["ФИО_подписанта_1_инициалы"] = abbreviate(fio1)
            row_data["ФИО_подписанта_2_инициалы"] = abbreviate(fio2)

            for doc_type in TEMPLATE_TYPES:
                template_subdir = os.path.join(TEMPLATE_DIR, doc_type)
                template_filename = f"{set_number}. {doc_type}.docx"
                template_path = os.path.join(template_subdir, template_filename)

                if os.path.isfile(template_path):
                    output_path = os.path.join(out_folder, f"{doc_type}.docx")
                    try:
                        process_docx(template_path, output_path, row_data)
                        total_files += 1
                        print(f"  [{contract_number}] {doc_type} — OK")
                    except Exception as e:
                        err = f"[{contract_number}] {doc_type}: {e}"
                        errors.append(err)
                        print(f"  [{contract_number}] {doc_type} — ОШИБКА: {e}")
                else:
                    print(f"  [{contract_number}] {doc_type} — ШАБЛОН НЕ НАЙДЕН: {template_filename}")

    reader.close()

    print(f"\n{'='*50}")
    print(f"Готово! Строк: {total_rows}, файлов: {total_files}")
    print(f"Папка: {os.path.abspath(OUTPUT_DIR)}")
    if errors:
        print(f"\nОшибки ({len(errors)}):")
        for e in errors:
            print(f"  • {e}")
    if sys.stdin.isatty():
        input("\nНажмите Enter для выхода...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"КРИТИЧЕСКАЯ ОШИБКА: {e}")
        traceback.print_exc()
        if sys.stdin.isatty():
            input("Нажмите Enter для выхода...")
