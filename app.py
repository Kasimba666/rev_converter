"""
Ревизские сказки — конвертер DOCX → XLSX
Этап 1: сканирование DOCX и формирование JSON-реестра.

Запуск:
    streamlit run app.py
"""

from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st
from docx import Document

REGISTRY_DIR = Path(__file__).resolve().parent / "registry"
REGISTRY_DIR.mkdir(parents=True, exist_ok=True)


def scan_docx_files(root_path: str) -> Dict[str, object]:
    """
    Сканирует общую папку и ищет .docx файлы исключительно в её подпапках.

    Возвращает словарь:
    {
        "root": str,
        "subfolders": list[str],
        "files": list[dict],       # [{subfolder, filename, path}, ...]
        "root_level_files": list[dict],  # файлы .docx прямо в корне (не в реестре)
    }
    """
    root = Path(root_path)

    result: Dict[str, object] = {
        "root": str(root.resolve()) if root.exists() else root_path,
        "subfolders": [],
        "files": [],
        "root_level_files": [],
    }

    if not root.exists() or not root.is_dir():
        return result

    # Файлы .docx прямо в корне — по правилу не учитываются
    for item in sorted(root.iterdir()):
        if item.is_file() and item.suffix.lower() == ".docx" and not item.name.startswith("~$"):
            result["root_level_files"].append(
                {"filename": item.name, "path": str(item.resolve())}
            )

    # Обход только прямых подпапок (docx размещаются исключительно в них)
    subfolders = sorted([p for p in root.iterdir() if p.is_dir()])
    result["subfolders"] = [p.name for p in subfolders]

    for subfolder in subfolders:
        for docx_file in sorted(subfolder.glob("*.docx")):
            if docx_file.name.startswith("~$"):
                continue
            result["files"].append(
                {
                    "subfolder": subfolder.name,
                    "filename": docx_file.name,
                    "path": str(docx_file.resolve()),
                }
            )

    return result


def extract_docx_to_json(file_path: str) -> Dict[str, object]:
    """
    Извлекает максимум информации из DOCX-файла:
    - метаданные файла (имя, путь, размер, даты)
    - свойства документа (заголовок, автор, дата создания, теги и т.д.)
    - параметры разделов (размер страницы, ориентация, поля)
    - все абзацы со стилями и выравниванием
    - все таблицы с заголовками и строками
    - текст заголовков и подвалов
    - количество изображений
    - определение ревизии и года
    """
    path = Path(file_path)
    stat = path.stat()
    try:
        doc = Document(str(path))
    except Exception as e:
        return {
            "file_info": {
                "name": path.name,
                "stem": path.stem,
                "subfolder": path.parent.name,
                "path": str(path.resolve()),
                "size_bytes": stat.st_size,
                "size_kb": round(stat.st_size / 1024, 2),
                "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                "created": datetime.fromtimestamp(stat.st_ctime).isoformat(),
            },
            "error": f"Не удалось прочитать файл: {str(e)}",
            "core_properties": {},
            "sections": [],
            "revision": "",
            "year": None,
            "paragraphs": [],
            "tables": [],
            "headers": [],
            "footers": [],
            "image_count": 0,
            "stats": {
                "paragraphs_count": 0,
                "tables_count": 0,
                "headers_count": 0,
                "footers_count": 0,
                "images_count": 0,
            },
        }

    # --- Метаданные файла ---
    file_info = {
        "name": path.name,
        "stem": path.stem,
        "subfolder": path.parent.name,
        "path": str(path.resolve()),
        "size_bytes": stat.st_size,
        "size_kb": round(stat.st_size / 1024, 2),
        "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
        "created": datetime.fromtimestamp(stat.st_ctime).isoformat(),
    }

    # --- Свойства документа (core properties) ---
    cp = doc.core_properties
    core_properties = {
        "title": cp.title or "",
        "author": cp.author or "",
        "subject": cp.subject or "",
        "keywords": cp.keywords or "",
        "comments": cp.comments or "",
        "category": cp.category or "",
        "created": str(cp.created) if cp.created else "",
        "modified": str(cp.modified) if cp.modified else "",
        "last_modified_by": cp.last_modified_by or "",
        "revision": cp.revision if cp.revision else 0,
        "version": cp.version or "",
    }

    # --- Разделы (размер страницы, поля) ---
    sections = []
    for s in doc.sections:
        sections.append({
            "page_width_cm": round(s.page_width.cm, 2) if s.page_width else 0,
            "page_height_cm": round(s.page_height.cm, 2) if s.page_height else 0,
            "orientation": str(s.orientation),
            "left_margin_cm": round(s.left_margin.cm, 2) if s.left_margin else 0,
            "right_margin_cm": round(s.right_margin.cm, 2) if s.right_margin else 0,
            "top_margin_cm": round(s.top_margin.cm, 2) if s.top_margin else 0,
            "bottom_margin_cm": round(s.bottom_margin.cm, 2) if s.bottom_margin else 0,
            "header_distance_cm": round(s.header_distance.cm, 2) if s.header_distance else 0,
            "footer_distance_cm": round(s.footer_distance.cm, 2) if s.footer_distance else 0,
        })

    # --- Абзацы со стилями ---
    paragraphs = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            paragraphs.append({
                "text": text,
                "style": p.style.name if p.style else "Normal",
                "alignment": str(p.alignment) if p.alignment else "left",
            })

    # --- Таблицы ---
    tables_out = []
    for t in doc.tables:
        table_data = {
            "rows_count": len(t.rows),
            "cols_count": len(t.columns),
            "headers": [cell.text.strip() for cell in t.rows[0].cells] if t.rows else [],
            "rows": [],
        }
        for row in t.rows[1:]:
            table_data["rows"].append([cell.text.strip() for cell in row.cells])
        tables_out.append(table_data)

    # --- Заголовки и подвалы ---
    headers_text = []
    footers_text = []
    for s in doc.sections:
        if s.header:
            h_text = "\n".join([p.text.strip() for p in s.header.paragraphs if p.text.strip()])
            if h_text:
                headers_text.append(h_text)
        if s.footer:
            f_text = "\n".join([p.text.strip() for p in s.footer.paragraphs if p.text.strip()])
            if f_text:
                footers_text.append(f_text)

    # --- Изображения ---
    image_count = 0
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            image_count += 1

    # --- Ревизия и год ---
    revision_text = ""
    year = None
    for p in paragraphs:
        if "РЕВИЗИЯ" in p["text"].upper():
            revision_text = p["text"]
            m = re.search(r"(\d{4})", revision_text)
            if m:
                year = int(m.group(1))
            break

    return {
        "file_info": file_info,
        "core_properties": core_properties,
        "sections": sections,
        "revision": revision_text,
        "year": year,
        "paragraphs": paragraphs,
        "tables": tables_out,
        "headers": headers_text,
        "footers": footers_text,
        "image_count": image_count,
        "stats": {
            "paragraphs_count": len(paragraphs),
            "tables_count": len(tables_out),
            "headers_count": len(headers_text),
            "footers_count": len(footers_text),
            "images_count": image_count,
        },
    }


def json_to_final_xlsx_rows(extracted: List[dict], revision_filter: str = None) -> List[List[str]]:
    """
    Превращает извлечённые JSON-данные в плоский список строк для XLSX.
    Одна строка = одна комбинация «населённый пункт + сословие».
    Населённый пункт дублируется для каждого сословия.

    Колонки:
    Ревизия | Год | Исследователь | Файл | № | Населённый пункт | Сословие | ДМП | ДЖП | Листы
    """
    rows = []
    header = [
        "Ревизия", "Год", "Исследователь", "Файл",
        "№", "Населённый пункт", "Сословие", "ДМП", "ДЖП", "Листы",
    ]
    rows.append(header)

    for d in extracted:
        revision = d.get("revision") or ""
        if revision_filter and revision != revision_filter:
            continue

        year = d.get("year") or ""
        subfolder = d.get("file_info", {}).get("subfolder", "")
        file_name = d.get("file_info", {}).get("name", "")

        tables = d.get("tables", [])
        if not tables:
            rows.append([revision, year, subfolder, file_name, "", "", "", "", "", ""])
            continue

        for table in tables:
            table_rows = table.get("rows", [])
            if not table_rows:
                continue

            # Определяем индекс колонки "Населённый пункт" и "Сословие" по заголовкам
            headers = table.get("headers", [])
            idx_point = -1
            idx_soslovie = -1
            idx_nomer = -1
            idx_dmp = -1
            idx_djp = -1
            idx_listy = -1

            for j, h in enumerate(headers):
                hl = h.lower()
                if "насел" in hl:
                    idx_point = j
                elif "сослов" in hl or "конфессион" in hl:
                    idx_soslovie = j
                elif hl.strip() == "№":
                    idx_nomer = j
                elif hl == "ДМП":
                    idx_dmp = j
                elif hl == "ДЖП":
                    idx_djp = j
                elif "лист" in hl:
                    idx_listy = j

            last_point = ""
            for row in table_rows:
                # Безопасное получение значения по индексу
                def get(idx):
                    return row[idx] if idx >= 0 and idx < len(row) else ""

                point = get(idx_point)
                if point.strip():
                    last_point = point  # запомним последний непустой населённый пункт
                else:
                    point = last_point  # продублируем

                soslovie = get(idx_soslovie)
                nomer = get(idx_nomer)
                dmp = get(idx_dmp)
                djp = get(idx_djp)
                listy = get(idx_listy)

                rows.append([
                    revision,
                    str(year) if year else "",
                    subfolder,
                    file_name,
                    nomer,
                    point,
                    soslovie,
                    dmp,
                    djp,
                    listy,
                ])

    return rows


def create_final_xlsx_by_revision(extracted: List[dict]) -> Dict[str, bytes]:
    """
    Создаёт отдельный XLSX-файл для каждой уникальной ревизии.
    Возвращает словарь: {имя_ревизии: bytes_xlsx}
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    from io import BytesIO

    # Собираем уникальные ревизии
    revisions = []
    seen = set()
    for d in extracted:
        rev = d.get("revision") or "Без ревизии"
        if rev not in seen:
            seen.add(rev)
            revisions.append(rev)

    result = {}
    for rev in revisions:
        rows = json_to_final_xlsx_rows(extracted, revision_filter=rev)

        wb = Workbook()
        ws = wb.active
        safe_name = rev[:31] if rev else "Без ревизии"
        # Убираем недопустимые символы для имени листа Excel
        safe_name = "".join(c for c in safe_name if c not in "\\/?*[]:")
        ws.title = safe_name if safe_name else "Ревизия"

        for r_idx, row_data in enumerate(rows, start=1):
            for c_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                if r_idx == 1:
                    cell.font = Font(bold=True)

        # Ширина колонок
        widths = [40, 8, 20, 28, 6, 30, 35, 10, 10, 25]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        result[rev] = buf.read()

    return result


def json_rows_to_xlsx(extracted: List[dict]) -> bytes:
    """
    Превращает список извлечённых DOCX-данных в плоскую XLSX-таблицу.
    Каждая строка каждой таблицы из DOCX — отдельная строка в XLSX.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from io import BytesIO

    wb = Workbook()
    ws = wb.active
    ws.title = "Все ревизии"

    headers = [
        "Файл",
        "Папка",
        "Ревизия",
        "Год",
        "№",
        "Населённый пункт",
        "Сословие",
        "ДМП",
        "ДЖП",
        "Листы",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for d in extracted:
        file_name = d.get("file", "")
        subfolder = d.get("subfolder", "")
        revision = d.get("revision", "")
        year = d.get("year", "")

        tables = d.get("tables", [])
        if not tables:
            ws.append([file_name, subfolder, revision, year, "", "", "", "", "", ""])
            continue

        for table in tables:
            rows = table.get("rows", [])
            for row in rows:
                vals = (row + [""] * 6)[:6]
                ws.append(
                    [file_name, subfolder, revision, year] + vals
                )

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def create_structured_xlsx(extracted: List[dict]) -> bytes:
    """
    Создаёт структурированный XLSX по типу ревизионных сказок:
    - Лист 0: Реестр всех файлов
    - Каждая ревизия — отдельный лист с таблицей данных
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from io import BytesIO
    from collections import OrderedDict

    wb = Workbook()

    # ---------- Лист 0: Реестр файлов ----------
    ws0 = wb.active
    ws0.title = "Реестр файлов"
    ws0.append(["Файл", "Папка", "Ревизия", "Год", "Таблиц", "Строк"])
    for cell in ws0[1]:
        cell.font = Font(bold=True)

    for d in extracted:
        first_table = d["tables"][0] if d["tables"] else {"headers": [], "rows": []}
        ws0.append([
            d.get("file", ""),
            d.get("subfolder", ""),
            d.get("revision", ""),
            d.get("year", ""),
            len(d.get("tables", [])),
            len(first_table.get("rows", [])),
        ])

    # ---------- Листы по ревизиям ----------
    by_revision: OrderedDict[str, List[dict]] = OrderedDict()
    for d in extracted:
        rev = d.get("revision") or "Без ревизии"
        by_revision.setdefault(rev, []).append(d)

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for rev_name, items in by_revision.items():
        sheet_name = rev_name[:31] if rev_name else "Данные"
        ws = wb.create_sheet(sheet_name)

        ws.merge_cells("A1:J1")
        ws["A1"] = rev_name
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center")

        headers = [
            "Файл", "Папка", "№", "Населённый пункт",
            "Сословие", "ДМП", "ДЖП", "Листы",
        ]
        ws.append([])
        ws.append(headers)
        for cell in ws[3]:
            cell.font = Font(bold=True)
            cell.border = border

        row_idx = 4
        for d in items:
            tables = d.get("tables", [])
            if not tables:
                ws.append([d.get("file", ""), d.get("subfolder", "")] + [""] * 6)
                row_idx += 1
                continue
            for table in tables:
                for row in table.get("rows", []):
                    vals = (row + [""] * 6)[:6]
                    ws.append([d.get("file", ""), d.get("subfolder", "")] + vals)
                    for col in range(1, 10):
                        cell = ws.cell(row=row_idx, column=col)
                        cell.border = border
                    row_idx += 1

        widths = [28, 20, 8, 30, 40, 10, 10, 25]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def build_registry_payload(scan_result: Dict[str, object]) -> Dict[str, object]:
    """Формирует payload JSON-реестра из результата сканирования."""
    files: List[dict] = scan_result["files"]  # type: ignore[assignment]
    return {
        "root_folder": scan_result["root"],
        "scanned_at": datetime.now().isoformat(),
        "total_files": len(files),
        "files": [
            {"subfolder": item["subfolder"], "filename": item["filename"]}
            for item in files
        ],
    }


def save_registry_to_disk(payload: Dict[str, object], file_name: str = "docx_registry.json") -> Path:
    """Сохраняет реестр в папку registry/ и возвращает путь к файлу."""
    file_path = REGISTRY_DIR / file_name
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return file_path


def main() -> None:
    st.set_page_config(
        page_title="Реестр DOCX — ревизские сказки",
        page_icon="📄",
        layout="wide",
    )

    st.title("📄 Ревизские сказки — конвертер DOCX → XLSX")
    st.caption("Этап 1: сканирование DOCX и формирование JSON-реестра")

    st.divider()

    # ---------- Папка на диске ----------
    st.subheader("📁 Папка с DOCX")

    if "folder_path" not in st.session_state:
        st.session_state["folder_path"] = ""

    col_path, col_browse = st.columns([5, 1])
    with col_path:
        folder_path = st.text_input(
            "Путь к папке",
            value=st.session_state.get("folder_path", ""),
            placeholder="например: D:\\Архив\\Ревизские сказки",
        )
        if folder_path:
            st.session_state["folder_path"] = folder_path

    with col_browse:
        st.write("")
        if st.button("📂 Выбрать папку (Windows)", use_container_width=True):
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            selected = filedialog.askdirectory(title="Выберите папку с DOCX")
            root.destroy()

            if selected:
                st.session_state["folder_path"] = selected
                st.rerun()

    scan_clicked = st.button(
        "🔍 Сканировать папку и составить реестр",
        type="primary",
        use_container_width=True,
    )

    if scan_clicked:
        folder_path = st.session_state.get("folder_path", "")
        if not folder_path:
            st.error("Укажите путь к папке.")
        else:
            scan_result = scan_docx_files(folder_path)
            st.session_state["scan_result"] = scan_result
            st.session_state["docx_registry"] = build_registry_payload(scan_result)

    if "scan_result" in st.session_state:
        st.divider()
        st.subheader("📋 Найденные DOCX-файлы")

        scan_result: Dict[str, object] = st.session_state["scan_result"]
        registry: Dict[str, object] = st.session_state["docx_registry"]

        files: List[dict] = scan_result["files"]  # type: ignore[assignment]
        root_level_files: List[dict] = scan_result["root_level_files"]  # type: ignore[assignment]

        if root_level_files:
            st.warning(
                f"⚠️ {len(root_level_files)} DOCX-файлов в корне не учитываются (только подпапки)."
            )

        if not files and not root_level_files:
            st.info("DOCX-файлы не найдены.")
        elif not files:
            st.warning("В подпапках DOCX-файлы не найдены.")
        else:
            st.success(
                f"✅ Найдено **{len(files)}** DOCX-файлов "
                f"в {len(scan_result['subfolders'])} подпапках."
            )

            # Извлекаем JSON из каждого DOCX
            if st.button("📄 Извлечь данные из всех DOCX в JSON", type="primary", use_container_width=True):
                extracted = []
                progress = st.progress(0)
                for i, item in enumerate(files):
                    data = extract_docx_to_json(item["path"])
                    extracted.append(data)
                    progress.progress((i + 1) / len(files))
                st.session_state["docx_json"] = extracted
                progress.empty()

            # Показываем извлечённые JSON по каждому файлу
            if "docx_json" in st.session_state:
                extracted: List[dict] = st.session_state["docx_json"]
                st.divider()
                st.subheader("📦 Извлечённые JSON по файлам")

                total_stats = {
                    "files": len(extracted),
                    "paragraphs": sum(d["stats"]["paragraphs_count"] for d in extracted),
                    "tables": sum(d["stats"]["tables_count"] for d in extracted),
                    "images": sum(d["stats"]["images_count"] for d in extracted),
                }
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Файлов", total_stats["files"])
                col2.metric("Абзацев", total_stats["paragraphs"])
                col3.metric("Таблиц", total_stats["tables"])
                col4.metric("Изображений", total_stats["images"])

                # Индивидуальные отчёты по файлам
                for i, d in enumerate(extracted):
                    fi = d["file_info"]

                    # Состояние развёрнутости карточки файла (сохраняется между перерисовками)
                    expander_key = f"expander_{i}"
                    if expander_key not in st.session_state:
                        st.session_state[expander_key] = False

                    with st.expander(
                        f"📄 {fi['name']} ({fi['subfolder']}) — {fi['size_kb']} КБ",
                        expanded=st.session_state[expander_key],
                    ):
                        # Метаданные файла
                        st.markdown(
                            f"**Ревизия:** {d['revision'] or '—'} | "
                            f"**Год:** {d['year'] or '—'} | "
                            f"**Абзацев:** {d['stats']['paragraphs_count']} | "
                            f"**Таблиц:** {d['stats']['tables_count']} | "
                            f"**Изображений:** {d['stats']['images_count']}"
                        )

                        # Состояние показа для этого файла
                        mode_key = f"mode_{i}"
                        if mode_key not in st.session_state:
                            st.session_state[mode_key] = None

                        # Три кнопки действий
                        b1, b2, b3 = st.columns(3)
                        with b1:
                            if st.button(
                                "📄 Показать JSON",
                                key=f"btn_json_{i}",
                                use_container_width=True,
                            ):
                                st.session_state[mode_key] = "json"
                                st.session_state[expander_key] = True
                                st.rerun()
                        with b2:
                            if st.button(
                                "📊 Показать таблицу",
                                key=f"btn_table_{i}",
                                use_container_width=True,
                            ):
                                st.session_state[mode_key] = "table"
                                st.session_state[expander_key] = True
                                st.rerun()
                        with b3:
                            if st.button(
                                "📈 Показать XLSX",
                                key=f"btn_xlsx_{i}",
                                use_container_width=True,
                            ):
                                st.session_state[mode_key] = "xlsx"
                                st.session_state[expander_key] = True
                                st.rerun()

                        st.divider()

                        mode = st.session_state[mode_key]

                        # ---------- РЕЖИМ JSON ----------
                        if mode == "json":
                            json_bytes = json.dumps(d, ensure_ascii=False, indent=2).encode("utf-8")
                            st.code(json_bytes.decode("utf-8"), language="json")
                            st.download_button(
                                label="⬇️ Скачать JSON",
                                data=json_bytes,
                                file_name=f"{Path(fi['name']).stem}.json",
                                mime="application/json",
                                use_container_width=True,
                                key=f"dl_json_{i}",
                            )

                        # ---------- РЕЖИМ ТАБЛИЦА ----------
                        elif mode == "table":
                            tables = d.get("tables", [])
                            if not tables:
                                st.info("В файле не найдено таблиц.")
                            else:
                                for ti, table in enumerate(tables):
                                    st.markdown(f"### Таблица {ti + 1}")
                                    headers = table.get("headers", [])
                                    rows = table.get("rows", [])
                                    if headers and rows:
                                        df = pd.DataFrame(rows, columns=headers)
                                        # Высота, чтобы таблица отображалась полностью без внутреннего скролла
                                        row_height = 35
                                        header_height = 40
                                        table_height = header_height + len(rows) * row_height
                                        st.dataframe(
                                            df,
                                            use_container_width=True,
                                            hide_index=True,
                                            height=table_height,
                                        )
                                    else:
                                        st.write("Пустая таблица")

                            json_bytes = json.dumps(d, ensure_ascii=False, indent=2).encode("utf-8")
                            st.download_button(
                                label="⬇️ Скачать JSON",
                                data=json_bytes,
                                file_name=f"{Path(fi['name']).stem}.json",
                                mime="application/json",
                                use_container_width=True,
                                key=f"dl_json_table_{i}",
                            )

                        # ---------- РЕЖИМ XLSX ----------
                        elif mode == "xlsx":
                            xlsx_key = f"single_xlsx_{i}"
                            if xlsx_key not in st.session_state:
                                single_xlsx = create_structured_xlsx([d])
                                st.session_state[xlsx_key] = single_xlsx

                            st.success("XLSX-файл готов к скачиванию")
                            st.download_button(
                                label="⬇️ Скачать XLSX",
                                data=st.session_state[xlsx_key],
                                file_name=f"{Path(fi['name']).stem}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key=f"dl_xlsx_{i}",
                            )

                        # ---------- НЕТ РЕЖИМА ----------
                        else:
                            json_bytes = json.dumps(d, ensure_ascii=False, indent=2).encode("utf-8")
                            preview = json_bytes.decode("utf-8")[:800]
                            st.code(preview, language="json")
                            if len(json_bytes) > 800:
                                st.caption(
                                    f"... и ещё {len(json_bytes) - 800} символов. "
                                    f"Нажмите кнопку выше для полного просмотра."
                                )

                # Скачать все JSON одним файлом
                json_all = {
                    "source_folder": scan_result["root"],
                    "extracted_at": datetime.now().isoformat(),
                    "total_files": len(extracted),
                    "files": extracted,
                }
                json_all_bytes = json.dumps(json_all, ensure_ascii=False, indent=2).encode("utf-8")
                st.download_button(
                    label="⬇️ Скачать все JSON в одном файле",
                    data=json_all_bytes,
                    file_name="docx_all_data.json",
                    mime="application/json",
                    use_container_width=True,
                )

                # ---------- Конвертация в конечные XLSX ----------
                st.divider()
                st.subheader("📊 Конечные XLSX по ревизиям")

                st.caption(
                    "Одна строка = один населённый пункт + одно сословие. "
                    "Населённый пункт дублируется для каждого сословия. "
                    "Один файл на каждую уникальную ревизию."
                )

                # Кнопка создания всех конечных файлов сразу
                if st.button("📦 Создать все конечные XLSX (по ревизиям)", type="primary", use_container_width=True):
                    final_files = create_final_xlsx_by_revision(extracted)
                    st.session_state["final_xlsx_files"] = final_files

                if "final_xlsx_files" in st.session_state:
                    final_files: Dict[str, bytes] = st.session_state["final_xlsx_files"]
                    st.success(f"Готово: {len(final_files)} файл(а) по ревизиям")

                    for rev_name, xlsx_bytes in final_files.items():
                        safe_rev = "".join(c for c in rev_name if c.isalnum() or c in (" ", "-", "(", ")", "_")).strip()
                        file_name = f"{safe_rev}.xlsx" if safe_rev else "ревизия.xlsx"
                        col_r, col_dl = st.columns([4, 1])
                        with col_r:
                            st.markdown(f"**{rev_name}**")
                        with col_dl:
                            st.download_button(
                                label="⬇️ XLSX",
                                data=xlsx_bytes,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key=f"dl_final_{rev_name[:20]}",
                            )

                # ---------- Старая общая XLSX (опционально) ----------
                with st.expander("Дополнительно: общая плоская таблица", expanded=False):
                    if st.button("📈 Создать общую XLSX-таблицу (все ревизии в одном файле)", use_container_width=True):
                        xlsx_bytes = json_rows_to_xlsx(extracted)
                        st.session_state["docx_xlsx"] = xlsx_bytes

                    if st.button("📋 Создать структурированный XLSX (листы по ревизиям)", use_container_width=True):
                        struct_xlsx = create_structured_xlsx(extracted)
                        st.session_state["docx_struct_xlsx"] = struct_xlsx

                    if "docx_xlsx" in st.session_state:
                        st.download_button(
                            label="⬇️ Скачать общий XLSX (все ревизии)",
                            data=st.session_state["docx_xlsx"],
                            file_name="docx_all_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    if "docx_struct_xlsx" in st.session_state:
                        st.download_button(
                            label="⬇️ Скачать структурированный XLSX (по ревизиям)",
                            data=st.session_state["docx_struct_xlsx"],
                            file_name="docx_by_revision.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )

            # Таблица файлов с кнопкой открытия
            st.markdown("### 📄 Файлы (нажмите на имя чтобы открыть DOCX)")
            for i, item in enumerate(files):
                file_path = item["path"]
                file_name = item["filename"]
                col_open, col_sub, col_path = st.columns([3, 2, 5])
                with col_open:
                    if st.button(f"📄 {file_name}", key=f"open_file_{i}", use_container_width=True):
                        os.startfile(file_path)
                        st.success(f"Открыт: {file_name}")
                with col_sub:
                    st.caption(item["subfolder"])
                with col_path:
                    st.caption(file_path)

            json_bytes = json.dumps(registry, ensure_ascii=False, indent=2).encode("utf-8")

            st.download_button(
                label="⬇️ Скачать docx_registry.json",
                data=json_bytes,
                file_name="docx_registry.json",
                mime="application/json",
                use_container_width=True,
            )

            if st.button("💾 Сохранить в папку registry/", use_container_width=True):
                saved_path = save_registry_to_disk(registry)
                st.success(f"Сохранено: {saved_path}")


if __name__ == "__main__":
    main()

