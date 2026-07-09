import hmac
import os
import re
import sqlite3
import unicodedata
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
)

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "uretim_kayitlari.db"
DEFAULT_EXCEL_PATH = BASE_DIR / "ayar_dosyasi.xlsx"

WORKER_NAMES = [
    "",
    "Abdulaziz Tilevmuradov",
    "Aykut Akili",
    "Azamat Remetov",
    "Köksal Atıl",
    "Mücahit Bozkurt",
    "Nezir Azim",
    "Oğuzhan Çevik",
    "Otabek Seytimbetov",
    "Sadettin Şeker",
    "Salih Kayar",
    "Sercan Kaygusuz",
    "Şükrü Atıl",
    "Yaşettin Baysal",
    "Yusuf Ay",
]


# -----------------------------
# VERİTABANI
# -----------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS work_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tarih TEXT NOT NULL,
            operator_ismi TEXT NOT NULL,
            proje TEXT,
            pos TEXT NOT NULL,
            kombinasyon_adi TEXT NOT NULL,
            calisma_tipi TEXT NOT NULL,
            calisma_saati REAL NOT NULL,
            neden TEXT,
            siparis_adedi INTEGER NOT NULL,
            saglam_tamamlanan INTEGER NOT NULL,
            fire_adedi INTEGER NOT NULL,
            uretim_yuku INTEGER NOT NULL,
            notlar TEXT,
            created_at TEXT NOT NULL
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS operation_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            urun_no INTEGER NOT NULL,
            operasyon_sirasi INTEGER NOT NULL,
            operasyon_adi TEXT NOT NULL,
            yapildi INTEGER NOT NULL,
            fire_var INTEGER NOT NULL DEFAULT 0,
            fire_operasyonu TEXT,
            fire_notu TEXT,
            FOREIGN KEY(session_id) REFERENCES work_sessions(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS update_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            guncelleme_tarihi TEXT NOT NULL,
            guncelleyen TEXT NOT NULL,
            ek_calisma_saati REAL NOT NULL DEFAULT 0,
            onceki_tamamlanan INTEGER NOT NULL,
            yeni_tamamlanan INTEGER NOT NULL,
            onceki_fire INTEGER NOT NULL,
            yeni_fire INTEGER NOT NULL,
            durum TEXT NOT NULL,
            aciklama TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(session_id) REFERENCES work_sessions(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS production_output_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            oc_no TEXT NOT NULL,
            project_name TEXT NOT NULL DEFAULT '',
            pos TEXT NOT NULL,
            requested_qty INTEGER NOT NULL,
            boy_mm REAL NOT NULL,
            en_mm REAL NOT NULL,
            unit_area_mm2 REAL NOT NULL,
            source_name TEXT,
            imported_at TEXT NOT NULL,
            UNIQUE(oc_no, pos)
        )
        """
    )

    cur.execute("PRAGMA table_info(production_output_items)")
    production_output_columns = {
        row[1] for row in cur.fetchall()
    }
    if "project_name" not in production_output_columns:
        cur.execute(
            "ALTER TABLE production_output_items "
            "ADD COLUMN project_name TEXT NOT NULL DEFAULT ''"
        )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS production_output_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id INTEGER NOT NULL,
            production_date TEXT NOT NULL,
            operator_name TEXT NOT NULL,
            produced_qty INTEGER NOT NULL,
            produced_area_mm2 REAL NOT NULL,
            note TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(item_id) REFERENCES production_output_items(id)
        )
        """
    )

    conn.commit()
    conn.close()


def insert_session(session_data, operation_rows):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO work_sessions (
            tarih, operator_ismi, proje, pos, kombinasyon_adi,
            calisma_tipi, calisma_saati, neden,
            siparis_adedi, saglam_tamamlanan, fire_adedi, uretim_yuku,
            notlar, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            session_data["tarih"],
            session_data["operator_ismi"],
            session_data["proje"],
            session_data["pos"],
            session_data["kombinasyon_adi"],
            session_data["calisma_tipi"],
            session_data["calisma_saati"],
            session_data["neden"],
            session_data["siparis_adedi"],
            session_data["saglam_tamamlanan"],
            session_data["fire_adedi"],
            session_data["uretim_yuku"],
            session_data["notlar"],
            datetime.now().isoformat(timespec="seconds"),
        ),
    )

    session_id = cur.lastrowid

    cur.executemany(
        """
        INSERT INTO operation_entries (
            session_id, urun_no, operasyon_sirasi, operasyon_adi,
            yapildi, fire_var, fire_operasyonu, fire_notu
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        [
            (
                session_id,
                row["urun_no"],
                row["operasyon_sirasi"],
                row["operasyon_adi"],
                int(row["yapildi"]),
                int(row["fire_var"]),
                row["fire_operasyonu"],
                row["fire_notu"],
            )
            for row in operation_rows
        ],
    )

    conn.commit()
    conn.close()
    return session_id




def production_status(siparis_adedi: int, tamamlanan_adet: int) -> str:
    """Sipariş ve tamamlanan adede göre okunabilir üretim durumu döndürür."""
    siparis = max(int(siparis_adedi or 0), 0)
    tamamlanan = max(int(tamamlanan_adet or 0), 0)
    if tamamlanan <= 0:
        return "Başlanmadı"
    if tamamlanan >= siparis and siparis > 0:
        return "Tamamlandı"
    return "Devam ediyor"


def update_session_progress(
    session_id: int,
    new_completed: int,
    new_fire: int,
    added_hours: float,
    updated_by: str,
    update_date: str,
    update_note: str,
    operation_rows: list,
):
    """Kaydın ilerlemesini ve operasyonlarını tek işlem içinde günceller."""
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT siparis_adedi, saglam_tamamlanan, fire_adedi, calisma_saati "
            "FROM work_sessions WHERE id = ?",
            (int(session_id),),
        )
        previous = cur.fetchone()
        if previous is None:
            raise ValueError("Güncellenecek kayıt bulunamadı.")

        siparis_adedi, previous_completed, previous_fire, previous_hours = previous
        new_load = int(siparis_adedi) + int(new_fire)
        new_total_hours = float(previous_hours or 0) + float(added_hours or 0)
        durum = production_status(int(siparis_adedi), int(new_completed))

        cur.execute(
            """
            UPDATE work_sessions
            SET saglam_tamamlanan = ?,
                fire_adedi = ?,
                uretim_yuku = ?,
                calisma_saati = ?
            WHERE id = ?
            """,
            (
                int(new_completed),
                int(new_fire),
                int(new_load),
                float(new_total_hours),
                int(session_id),
            ),
        )

        cur.execute(
            "DELETE FROM operation_entries WHERE session_id = ?",
            (int(session_id),),
        )
        cur.executemany(
            """
            INSERT INTO operation_entries (
                session_id, urun_no, operasyon_sirasi, operasyon_adi,
                yapildi, fire_var, fire_operasyonu, fire_notu
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    int(session_id),
                    int(row["urun_no"]),
                    int(row["operasyon_sirasi"]),
                    str(row["operasyon_adi"]),
                    int(bool(row["yapildi"])),
                    int(bool(row["fire_var"])),
                    str(row.get("fire_operasyonu", "") or ""),
                    str(row.get("fire_notu", "") or ""),
                )
                for row in operation_rows
            ],
        )

        cur.execute(
            """
            INSERT INTO update_history (
                session_id, guncelleme_tarihi, guncelleyen,
                ek_calisma_saati, onceki_tamamlanan, yeni_tamamlanan,
                onceki_fire, yeni_fire, durum, aciklama, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(session_id),
                str(update_date),
                updated_by.strip(),
                float(added_hours or 0),
                int(previous_completed),
                int(new_completed),
                int(previous_fire),
                int(new_fire),
                durum,
                update_note.strip(),
                datetime.now().isoformat(timespec="seconds"),
            ),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_update_history(session_id: int) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql_query(
            """
            SELECT
                id,
                guncelleme_tarihi,
                guncelleyen,
                ek_calisma_saati,
                onceki_tamamlanan,
                yeni_tamamlanan,
                onceki_fire,
                yeni_fire,
                durum,
                aciklama,
                created_at
            FROM update_history
            WHERE session_id = ?
            ORDER BY id DESC
            """,
            conn,
            params=(int(session_id),),
        )
    finally:
        conn.close()


def delete_session(session_id: int):
    """Seçilen ana kaydı ve ona bağlı operasyon detaylarını siler."""
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM update_history WHERE session_id = ?",
            (int(session_id),),
        )
        cur.execute(
            "DELETE FROM operation_entries WHERE session_id = ?",
            (int(session_id),),
        )
        cur.execute(
            "DELETE FROM work_sessions WHERE id = ?",
            (int(session_id),),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()



# -----------------------------
# EXCEL AYAR DOSYASI OKUMA
# -----------------------------
def clean_text(x):
    if pd.isna(x):
        return None
    text = str(x).strip()
    return text if text else None


@st.cache_data(show_spinner=False)
def load_config_from_excel_bytes(file_bytes: bytes):
    """
    Excel mantığı:
    - Sayfa1: çalışma nedenleri / olasılıklar
    - Sayfa2: B sütunu uygulamada gösterilmez.
              C-L arası operasyon adımları olarak alınır.
    """
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet1 = pd.read_excel(xls, sheet_name="Sayfa1", header=None)
    sheet2 = pd.read_excel(xls, sheet_name="Sayfa2", header=None)

    ignored = {
        "girdiler", "tarih", "operatör ismi", "çalıştığı proje",
        "ürettiği sağlam", "üretiği fireli adet", "çalışma",
        "tam çalışma", "evet", "hayır", "adet", "m2", "mtül",
        "default", "saat gir"
    }

    reasons = []
    if sheet1.shape[1] > 1:
        for value in sheet1.iloc[:, 1].dropna().tolist():
            text = clean_text(value)
            if not text:
                continue
            normalized = text.lower().strip()
            if normalized not in ignored and text not in reasons:
                reasons.append(text)

    combinations = []
    operation_area = sheet2.iloc[:, 2:12]  # C-L

    for _, row in operation_area.iterrows():
        ops = [clean_text(v) for v in row.tolist()]
        ops = [v for v in ops if v is not None]

        if not ops:
            continue

        if all(str(v).replace(".0", "").isdigit() for v in ops):
            continue

        if ops not in [c["operasyonlar"] for c in combinations]:
            combinations.append(
                {
                    "ad": f"Kombinasyon {len(combinations) + 1}",
                    "operasyonlar": ops,
                }
            )

    if not reasons:
        reasons = ["Yarı zamanlı çalışma", "Bakım", "İzinli", "Mazeret", "Hastalık", "Eğitim", "Boş", "Diğer"]

    if not combinations:
        combinations = [
            {
                "ad": "Varsayılan Kombinasyon",
                "operasyonlar": ["laser", "abkant", "çıta", "kaynak", "taşlama", "boya", "paketleme", "sevkiyat"],
            }
        ]

    return reasons, combinations


def load_config(uploaded_excel):
    if uploaded_excel is not None:
        return load_config_from_excel_bytes(uploaded_excel.getvalue()), "Yüklenen Excel dosyası"

    if DEFAULT_EXCEL_PATH.exists():
        return load_config_from_excel_bytes(DEFAULT_EXCEL_PATH.read_bytes()), f"Varsayılan ayar dosyası: {DEFAULT_EXCEL_PATH.name}"

    return ([], []), "Excel ayar dosyası bulunamadı"



# -----------------------------
# ÜRETİM ÇIKTISI DOSYASI VE KAYITLARI
# -----------------------------
def _normalize_header(value) -> str:
    if pd.isna(value):
        return ""
    normalized = unicodedata.normalize("NFKD", str(value))
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]", "", normalized.lower())


def _to_float_tr(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    raw = str(value).strip().replace(" ", "")
    if not raw:
        return None
    if "," in raw and "." in raw:
        raw = raw.replace(".", "").replace(",", ".")
    else:
        raw = raw.replace(",", ".")
    match = re.search(r"-?\d+(?:\.\d+)?", raw)
    return float(match.group(0)) if match else None


def _to_int_value(value):
    parsed = _to_float_tr(value)
    return int(round(parsed)) if parsed is not None else None


def _extract_oc_no(file_name: str, text_content: str = "") -> str:
    for source in (file_name or "", text_content or ""):
        match = re.search(r"(?<!\d)(\d{5,})(?!\d)", source)
        if match:
            return match.group(1)
    return ""


def _clean_project_name(value, oc_no: str = "") -> str:
    if value is None or pd.isna(value):
        return ""
    project = re.sub(r"\s+", " ", str(value)).strip(" -_")
    if not project or project.lower() == "nan":
        return ""
    if oc_no:
        project = re.sub(
            rf"^\s*{re.escape(str(oc_no))}\s*[-_:]*\s*",
            "",
            project,
        ).strip()
    return project


def _extract_project_name(
    file_name: str,
    text_content: str = "",
    oc_no: str = "",
) -> str:
    text = re.sub(r"\s+", " ", text_content or "").strip()

    if oc_no and text:
        patterns = [
            rf"(?<!\d){re.escape(str(oc_no))}(?!\d)\s+(.+?)\s+Q\s*[-=]*>\s*R\b",
            rf"(?<!\d){re.escape(str(oc_no))}(?!\d)\s+(.+?)\s+\d+(?:[xX]\d+)+",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                project = _clean_project_name(match.group(1), oc_no)
                if project:
                    return project

    stem = Path(file_name or "").stem
    stem = re.sub(
        r"(?i)\b(üretim|uretim|çıktısı|ciktisi|output|liste|listesi)\b",
        " ",
        stem,
    )
    if oc_no:
        stem = re.sub(rf"(?<!\d){re.escape(str(oc_no))}(?!\d)", " ", stem)
    stem = _clean_project_name(stem, oc_no)
    return stem or (f"OC {oc_no}" if oc_no else "İsimsiz Proje")


def _pos_label(value) -> str:
    if value is None or pd.isna(value):
        return ""
    raw = str(value).strip().upper().replace(" ", "")
    match = re.search(r"(\d+)", raw)
    if not match:
        return ""
    return f"POS{int(match.group(1))}"


def _finalize_output_rows(rows: list[dict]) -> pd.DataFrame:
    columns = ["oc_no", "project_name", "pos", "requested_qty", "boy_mm", "en_mm", "unit_area_mm2"]
    if not rows:
        return pd.DataFrame(columns=columns)
    result = pd.DataFrame(rows)
    result = result.dropna(subset=["oc_no", "pos", "requested_qty", "boy_mm", "en_mm"])
    result["oc_no"] = result["oc_no"].astype(str).str.strip()
    if "project_name" not in result.columns:
        result["project_name"] = ""
    result["project_name"] = result.apply(
        lambda row: _clean_project_name(row.get("project_name"), row["oc_no"])
        or f"OC {row['oc_no']}",
        axis=1,
    )
    result["pos"] = result["pos"].map(_pos_label)
    result["requested_qty"] = pd.to_numeric(result["requested_qty"], errors="coerce").fillna(0).astype(int)
    result["boy_mm"] = pd.to_numeric(result["boy_mm"], errors="coerce")
    result["en_mm"] = pd.to_numeric(result["en_mm"], errors="coerce")
    result = result[
        result["oc_no"].ne("")
        & result["pos"].ne("")
        & result["requested_qty"].gt(0)
        & result["boy_mm"].gt(0)
        & result["en_mm"].gt(0)
    ].copy()
    result["unit_area_mm2"] = result["boy_mm"] * result["en_mm"]
    result["_pos_no"] = result["pos"].str.extract(r"(\d+)")[0].astype(int)
    result = (
        result.sort_values(["oc_no", "_pos_no"])
        .drop_duplicates(["oc_no", "pos"], keep="last")
        .drop(columns=["_pos_no"])
        .reset_index(drop=True)
    )
    return result[columns]


@st.cache_data(show_spinner=False)
def parse_production_output_pdf(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    reader = PdfReader(BytesIO(file_bytes))
    text_content = "\n".join(page.extract_text() or "" for page in reader.pages)
    oc_no = _extract_oc_no(file_name, text_content)
    if not oc_no:
        raise ValueError("PDF içinde OC No bulunamadı.")
    project_name = _extract_project_name(file_name, text_content, oc_no)

    row_pattern = re.compile(
        r"^\s*(\d+)\s+(\d+)\s+(\d+)\s+"
        r"(\d+(?:[.,]\d+)?)\s+(\d+(?:[.,]\d+)?)",
        re.MULTILINE,
    )
    rows = []
    for match in row_pattern.finditer(text_content):
        _, pos_value, qty_value, boy_value, en_value = match.groups()
        rows.append({
            "oc_no": oc_no,
            "project_name": project_name,
            "pos": _pos_label(pos_value),
            "requested_qty": int(qty_value),
            "boy_mm": _to_float_tr(boy_value),
            "en_mm": _to_float_tr(en_value),
        })

    result = _finalize_output_rows(rows)
    if result.empty:
        raise ValueError("PDF tablosundaki POS, Adet, Boy ve En satırları okunamadı.")
    return result


def _find_excel_column(columns, aliases):
    normalized_aliases = {_normalize_header(alias) for alias in aliases}
    for column in columns:
        normalized = _normalize_header(column)
        if normalized in normalized_aliases:
            return column
    return None


@st.cache_data(show_spinner=False)
def parse_production_output_excel(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    workbook = pd.ExcelFile(BytesIO(file_bytes))
    fallback_oc = _extract_oc_no(file_name)
    rows = []

    for sheet_name in workbook.sheet_names:
        raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None)
        header_index = None
        for index in range(min(len(raw), 30)):
            normalized_values = {_normalize_header(value) for value in raw.iloc[index].tolist()}
            if {"pos", "adet", "boy", "en"}.issubset(normalized_values):
                header_index = index
                break
        if header_index is None:
            continue

        headers = [str(value).strip() if not pd.isna(value) else f"bos_{i}" for i, value in enumerate(raw.iloc[header_index])]
        data = raw.iloc[header_index + 1:].copy()
        data.columns = headers

        pos_col = _find_excel_column(data.columns, ["POS"])
        qty_col = _find_excel_column(data.columns, ["Adet", "Miktar"])
        boy_col = _find_excel_column(data.columns, ["Boy", "Boy mm"])
        en_col = _find_excel_column(data.columns, ["En", "En mm"])
        oc_col = _find_excel_column(data.columns, ["OC No", "OC", "OC Numarası"])
        project_col = _find_excel_column(
            data.columns,
            ["Proje Adı", "Proje", "Project", "Project Name"],
        )
        if not all([pos_col, qty_col, boy_col, en_col]):
            continue

        oc_values = data[oc_col].ffill().bfill() if oc_col else pd.Series([fallback_oc] * len(data), index=data.index)
        fallback_project = _extract_project_name(file_name, "", fallback_oc)
        project_values = (
            data[project_col].ffill().bfill()
            if project_col
            else pd.Series([fallback_project] * len(data), index=data.index)
        )
        for index, row in data.iterrows():
            pos_value = _pos_label(row[pos_col])
            qty_value = _to_int_value(row[qty_col])
            boy_value = _to_float_tr(row[boy_col])
            en_value = _to_float_tr(row[en_col])
            oc_value = str(oc_values.loc[index]).strip() if index in oc_values.index and not pd.isna(oc_values.loc[index]) else fallback_oc
            if not oc_value or oc_value.lower() == "nan":
                oc_value = fallback_oc
            project_value = (
                _clean_project_name(project_values.loc[index], oc_value)
                if index in project_values.index
                else fallback_project
            )
            if not project_value:
                project_value = fallback_project or f"OC {oc_value}"
            if pos_value and qty_value and boy_value and en_value and oc_value:
                rows.append({
                    "oc_no": oc_value,
                    "project_name": project_value,
                    "pos": pos_value,
                    "requested_qty": qty_value,
                    "boy_mm": boy_value,
                    "en_mm": en_value,
                })

    result = _finalize_output_rows(rows)
    if result.empty:
        raise ValueError("Excel içinde OC No, POS, Adet, Boy ve En sütunları okunamadı.")
    return result


def parse_production_output_file(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    suffix = Path(file_name).suffix.lower()
    if suffix == ".pdf":
        return parse_production_output_pdf(file_bytes, file_name)
    if suffix in {".xlsx", ".xls"}:
        return parse_production_output_excel(file_bytes, file_name)
    raise ValueError("Sadece PDF, XLSX veya XLS üretim çıktısı yüklenebilir.")


def import_production_output_items(rows: pd.DataFrame, source_name: str) -> int:
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        imported_at = datetime.now().isoformat(timespec="seconds")
        for _, row in rows.iterrows():
            cur.execute(
                """
                INSERT INTO production_output_items (
                    oc_no, project_name, pos, requested_qty, boy_mm, en_mm,
                    unit_area_mm2, source_name, imported_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(oc_no, pos) DO UPDATE SET
                    project_name = excluded.project_name,
                    requested_qty = excluded.requested_qty,
                    boy_mm = excluded.boy_mm,
                    en_mm = excluded.en_mm,
                    unit_area_mm2 = excluded.unit_area_mm2,
                    source_name = excluded.source_name,
                    imported_at = excluded.imported_at
                """,
                (
                    str(row["oc_no"]),
                    str(row.get("project_name") or f"OC {row['oc_no']}"),
                    str(row["pos"]),
                    int(row["requested_qty"]),
                    float(row["boy_mm"]),
                    float(row["en_mm"]),
                    float(row["unit_area_mm2"]),
                    source_name,
                    imported_at,
                ),
            )
        conn.commit()
        return len(rows)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_production_output_summary(oc_no: str | None = None) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        query = """
            SELECT
                i.id AS item_id,
                i.oc_no,
                COALESCE(NULLIF(i.project_name, ''), 'OC ' || i.oc_no) AS project_name,
                i.pos,
                i.requested_qty,
                i.boy_mm,
                i.en_mm,
                i.unit_area_mm2,
                i.source_name,
                i.imported_at,
                COALESCE(SUM(l.produced_qty), 0) AS produced_qty
            FROM production_output_items i
            LEFT JOIN production_output_logs l ON l.item_id = i.id
        """
        params = ()
        if oc_no:
            query += " WHERE i.oc_no = ?"
            params = (str(oc_no),)
        query += " GROUP BY i.id ORDER BY i.oc_no, i.pos"
        result = pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()

    if result.empty:
        return result
    result["requested_qty"] = result["requested_qty"].astype(int)
    result["produced_qty"] = result["produced_qty"].astype(int)
    result["remaining_qty"] = (result["requested_qty"] - result["produced_qty"]).clip(lower=0)
    result["requested_area_mm2"] = result["unit_area_mm2"] * result["requested_qty"]
    result["produced_area_mm2"] = result["unit_area_mm2"] * result["produced_qty"]
    result["remaining_area_mm2"] = result["unit_area_mm2"] * result["remaining_qty"]
    result["_pos_no"] = result["pos"].str.extract(r"(\d+)")[0].fillna(0).astype(int)
    return result.sort_values(["oc_no", "_pos_no"]).drop(columns=["_pos_no"]).reset_index(drop=True)



def delete_production_output_oc(oc_no: str) -> dict:
    """Bir OC'ye ait yüklenen çıktı satırlarını ve günlük çıktı loglarını siler."""
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id FROM production_output_items WHERE oc_no = ?",
            (str(oc_no),),
        )
        item_ids = [int(row[0]) for row in cur.fetchall()]

        deleted_logs = 0
        if item_ids:
            placeholders = ",".join("?" for _ in item_ids)
            cur.execute(
                f"DELETE FROM production_output_logs "
                f"WHERE item_id IN ({placeholders})",
                item_ids,
            )
            deleted_logs = int(cur.rowcount or 0)

        cur.execute(
            "DELETE FROM production_output_items WHERE oc_no = ?",
            (str(oc_no),),
        )
        deleted_items = int(cur.rowcount or 0)
        conn.commit()

        return {
            "deleted_items": deleted_items,
            "deleted_logs": deleted_logs,
        }
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def sidebar_production_output_manager():
    """Yönetici için çoklu çıktı yükleme ve OC silme alanı."""
    st.subheader("Üretim Çıktısı")

    uploader_version = int(
        st.session_state.get("production_sidebar_uploader_version", 0)
    )
    uploaded_outputs = st.file_uploader(
        "PDF / Excel yükle",
        type=["pdf", "xlsx", "xls"],
        accept_multiple_files=True,
        key=f"production_sidebar_upload_{uploader_version}",
        help="Aynı anda birden fazla üretim çıktısı seçebilirsin.",
    )

    parsed_files = []
    if uploaded_outputs:
        for uploaded_output in uploaded_outputs:
            try:
                parsed = parse_production_output_file(
                    uploaded_output.getvalue(),
                    uploaded_output.name,
                )
                parsed_files.append((uploaded_output.name, parsed))

                oc_list = sorted(
                    parsed["oc_no"].dropna().astype(str).unique().tolist()
                )
                st.caption(
                    f"{uploaded_output.name}: "
                    f"{', '.join('OC ' + oc for oc in oc_list)} · "
                    f"{len(parsed)} POS"
                )
            except Exception as exc:
                st.error(f"{uploaded_output.name} okunamadı: {exc}")

        if parsed_files and st.button(
            "Dosyaları Sisteme Aktar",
            type="primary",
            use_container_width=True,
            key="sidebar_import_outputs",
        ):
            imported_count = 0
            for source_name, parsed in parsed_files:
                imported_count += import_production_output_items(
                    parsed,
                    source_name,
                )

            st.session_state.production_sidebar_uploader_version = (
                uploader_version + 1
            )
            st.session_state.production_sidebar_message = (
                f"{len(parsed_files)} dosyadan {imported_count} POS "
                "satırı aktarıldı."
            )
            st.rerun()

    sidebar_message = st.session_state.pop(
        "production_sidebar_message",
        None,
    )
    if sidebar_message:
        st.success(sidebar_message)

    all_items = get_production_output_summary()
    if all_items.empty:
        st.caption("Henüz yüklenmiş OC yok.")
        return

    st.divider()
    st.markdown("**Yüklenen OC’leri Yönet**")

    oc_summary = (
        all_items.groupby(["oc_no", "project_name"], as_index=False)
        .agg(
            pos_sayisi=("item_id", "count"),
            istenen_adet=("requested_qty", "sum"),
            istenen_alan_mm2=("requested_area_mm2", "sum"),
        )
        .sort_values("oc_no")
    )

    labels = oc_summary.apply(
        lambda row: (
            f"OC {row['oc_no']} · {row['project_name']} "
            f"({int(row['pos_sayisi'])} POS)"
        ),
        axis=1,
    ).tolist()
    label_to_oc = dict(
        zip(labels, oc_summary["oc_no"].astype(str))
    )

    selected_label = st.selectbox(
        "Silinecek OC",
        labels,
        key="sidebar_delete_oc_select",
    )
    selected_oc = label_to_oc[selected_label]
    selected_summary = oc_summary[
        oc_summary["oc_no"].astype(str) == selected_oc
    ].iloc[0]

    st.caption(
        f"İstenen: {int(selected_summary['istenen_adet'])} adet · "
        f"{_format_mm2(selected_summary['istenen_alan_mm2'])} mm²"
    )

    delete_confirm = st.checkbox(
        "Bu OC’yi silmek istediğimi onaylıyorum",
        key=f"sidebar_delete_oc_confirm_{selected_oc}",
    )

    if st.button(
        "Seçili OC’yi Sil",
        disabled=not delete_confirm,
        use_container_width=True,
        key=f"sidebar_delete_oc_button_{selected_oc}",
    ):
        result = delete_production_output_oc(selected_oc)
        st.session_state.production_sidebar_message = (
            f"OC {selected_oc} silindi. "
            f"{result['deleted_items']} POS satırı ve "
            f"{result['deleted_logs']} üretim hareketi kaldırıldı."
        )
        st.rerun()



def save_production_output_entries(
    entries: pd.DataFrame,
    production_date: str,
    operator_name: str,
    note: str,
) -> int:
    quantity_column = (
        "Bugün Üretilen"
        if "Bugün Üretilen" in entries.columns
        else "Bu İşlemde Üretilen"
    )
    selected = entries[
        pd.to_numeric(entries[quantity_column], errors="coerce")
        .fillna(0)
        .gt(0)
    ].copy()
    if selected.empty:
        return 0

    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        created_at = datetime.now().isoformat(timespec="seconds")
        for _, row in selected.iterrows():
            quantity = int(row[quantity_column])
            cur.execute(
                "SELECT unit_area_mm2 FROM production_output_items WHERE id = ?",
                (int(row["item_id"]),),
            )
            result = cur.fetchone()
            if result is None:
                raise ValueError("Üretim çıktısı satırı bulunamadı.")
            unit_area = float(result[0])
            cur.execute(
                """
                INSERT INTO production_output_logs (
                    item_id, production_date, operator_name,
                    produced_qty, produced_area_mm2, note, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    int(row["item_id"]),
                    production_date,
                    operator_name,
                    quantity,
                    quantity * unit_area,
                    note,
                    created_at,
                ),
            )
        conn.commit()
        return len(selected)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_production_output_history(oc_no: str) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql_query(
            """
            SELECT
                l.production_date AS tarih,
                l.operator_name AS operator_ismi,
                COALESCE(NULLIF(i.project_name, ''), 'OC ' || i.oc_no) AS Proje,
                i.oc_no AS OC,
                i.pos AS POS,
                l.produced_qty AS üretilen_adet,
                l.produced_area_mm2 AS üretilen_alan_mm2,
                l.note AS notlar,
                l.created_at AS kayıt_zamanı
            FROM production_output_logs l
            INNER JOIN production_output_items i ON i.id = l.item_id
            WHERE i.oc_no = ?
            ORDER BY l.id DESC
            """,
            conn,
            params=(str(oc_no),),
        )
    finally:
        conn.close()


def _format_mm2(value) -> str:
    numeric = float(value or 0)
    return f"{numeric:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _rename_oc_column(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={"proje": "OC", "Proje": "OC"})


# -----------------------------
# OPERASYON SEÇİM YARDIMCILARI
# -----------------------------
def set_all_operations_for_unit(urun_no: int, operation_count: int):
    """Bir ürünün tüm operasyon kutularını tek seferde seçer veya kaldırır."""
    selected = bool(st.session_state.get(f"select_all_{urun_no}", False))
    for idx in range(1, operation_count + 1):
        st.session_state[f"op_{urun_no}_{idx}"] = selected


def set_all_operations_for_all_units(unit_count: int, operation_count: int):
    """Ekrandaki bütün ürünlerin bütün operasyonlarını tek seferde seçer veya kaldırır."""
    selected = bool(st.session_state.get("select_all_all_units", False))
    for urun_no in range(1, unit_count + 1):
        st.session_state[f"select_all_{urun_no}"] = selected
        for idx in range(1, operation_count + 1):
            st.session_state[f"op_{urun_no}_{idx}"] = selected


def set_all_edit_operations_for_unit(
    session_id: int,
    urun_no: int,
    operation_count: int,
):
    selected = bool(
        st.session_state.get(f"edit_select_all_{session_id}_{urun_no}", False)
    )
    for idx in range(1, operation_count + 1):
        st.session_state[f"edit_op_{session_id}_{urun_no}_{idx}"] = selected


def set_all_edit_operations_for_all_units(
    session_id: int,
    unit_count: int,
    operation_count: int,
):
    selected = bool(
        st.session_state.get(f"edit_select_all_all_{session_id}", False)
    )
    for urun_no in range(1, unit_count + 1):
        st.session_state[f"edit_select_all_{session_id}_{urun_no}"] = selected
        for idx in range(1, operation_count + 1):
            st.session_state[f"edit_op_{session_id}_{urun_no}_{idx}"] = selected


def clear_edit_widget_state(session_id: int):
    prefixes = (
        f"edit_op_{session_id}_",
        f"edit_fire_{session_id}_",
        f"edit_fireop_{session_id}_",
        f"edit_firenote_{session_id}_",
        f"edit_select_all_{session_id}_",
        f"edit_select_all_all_{session_id}",
        f"edit_date_{session_id}",
        f"edit_by_{session_id}",
        f"edit_hours_{session_id}",
        f"edit_completed_{session_id}",
        f"edit_fire_count_{session_id}",
        f"edit_note_{session_id}",
        f"update_progress_button_{session_id}",
    )
    for key in list(st.session_state.keys()):
        if any(str(key).startswith(prefix) for prefix in prefixes):
            del st.session_state[key]


# -----------------------------
# YETKİLENDİRME
# -----------------------------
def get_manager_password() -> str:
    """Yönetici şifresini Streamlit Secrets veya ortam değişkeninden okur."""
    try:
        password = st.secrets.get("MANAGER_PASSWORD", "")
    except Exception:
        password = ""

    if not password:
        password = os.getenv("MANAGER_PASSWORD", "")

    return str(password)


def init_auth_state():
    if "manager_authenticated" not in st.session_state:
        st.session_state.manager_authenticated = False
    if "show_manager_login" not in st.session_state:
        st.session_state.show_manager_login = False


def manager_login_panel():
    """Kenar çubuğunda yönetici girişini veya çıkışını gösterir."""
    if st.session_state.manager_authenticated:
        st.success("Yönetici görünümü açık")
        if st.button("Çalışan Görünümüne Dön", use_container_width=True):
            st.session_state.manager_authenticated = False
            st.session_state.show_manager_login = False
            st.rerun()
        return

    st.info("Çalışan görünümü")

    if st.button("Yönetici Girişi", use_container_width=True):
        st.session_state.show_manager_login = True

    if not st.session_state.show_manager_login:
        return

    configured_password = get_manager_password()
    if not configured_password:
        st.error("Yönetici şifresi henüz Streamlit Secrets bölümünde tanımlanmamış.")
        st.code('MANAGER_PASSWORD = "guclu-bir-sifre"', language="toml")
        if st.button("Giriş ekranını kapat", use_container_width=True):
            st.session_state.show_manager_login = False
            st.rerun()
        return

    entered_password = st.text_input(
        "Yönetici şifresi",
        type="password",
        key="manager_password_input",
    )

    login_col, cancel_col = st.columns(2)
    with login_col:
        login_clicked = st.button("Giriş Yap", type="primary", use_container_width=True)
    with cancel_col:
        cancel_clicked = st.button("Vazgeç", use_container_width=True)

    if cancel_clicked:
        st.session_state.show_manager_login = False
        st.session_state.pop("manager_password_input", None)
        st.rerun()

    if login_clicked:
        if hmac.compare_digest(str(entered_password), configured_password):
            st.session_state.manager_authenticated = True
            st.session_state.show_manager_login = False
            st.session_state.pop("manager_password_input", None)
            st.rerun()
        else:
            st.error("Yönetici şifresi yanlış.")



def get_overtime_summary(target_date: str | None = None) -> pd.DataFrame:
    """Mesaili çalışma kayıtlarını işçi ve gün bazında birleştirir."""
    conn = sqlite3.connect(DB_PATH)
    try:
        query = """
            SELECT
                tarih,
                operator_ismi,
                ROUND(SUM(calisma_saati), 2) AS toplam_saat,
                GROUP_CONCAT(DISTINCT proje) AS OC,
                GROUP_CONCAT(DISTINCT pos) AS POS,
                SUM(saglam_tamamlanan) AS uretilen_adet,
                SUM(fire_adedi) AS fire_adedi
            FROM work_sessions
            WHERE calisma_tipi = 'Mesaili'
        """
        params = ()
        if target_date:
            query += " AND tarih = ?"
            params = (str(target_date),)
        query += """
            GROUP BY tarih, operator_ismi
            ORDER BY tarih DESC, operator_ismi
        """
        return pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()


def manager_overtime_sidebar_alert():
    """Yöneticiye bugünkü mesaili çalışma kayıtlarını kenar menüde gösterir."""
    overtime = get_overtime_summary(str(date.today()))
    if overtime.empty:
        return

    st.warning(f"Bugün {len(overtime)} mesai uyarısı var.")
    with st.expander("Mesai uyarılarını göster", expanded=False):
        display = overtime.rename(
            columns={
                "tarih": "Tarih",
                "operator_ismi": "İşçi",
                "toplam_saat": "Saat",
                "uretilen_adet": "Üretilen",
                "fire_adedi": "Fire",
            }
        )
        st.dataframe(display, use_container_width=True, hide_index=True)


def set_session_value(key: str, value):
    """Bir butona basıldığında ilgili giriş alanını doldurur."""
    st.session_state[key] = value


def build_quick_operation_rows(
    produced_qty: int,
    fire_qty: int,
    operations: list[str],
    completed_operations: list[str],
    fire_operation: str,
    fire_note: str,
) -> list[dict]:
    """Çoklu POS hızlı girişinde operasyon satırlarını üretir."""
    total_units = int(produced_qty) + int(fire_qty)
    rows = []
    completed = set(completed_operations)
    first_fire_unit = int(produced_qty) + 1

    for unit_no in range(1, total_units + 1):
        is_fire = unit_no >= first_fire_unit and int(fire_qty) > 0
        for operation_index, operation_name in enumerate(operations, start=1):
            rows.append(
                {
                    "urun_no": unit_no,
                    "operasyon_sirasi": operation_index,
                    "operasyon_adi": operation_name,
                    "yapildi": operation_name in completed,
                    "fire_var": is_fire,
                    "fire_operasyonu": fire_operation if is_fire else "",
                    "fire_notu": fire_note if is_fire else "",
                }
            )
    return rows



# -----------------------------
# UYGULAMA
# -----------------------------
def main():
    st.set_page_config(page_title="POS Üretim Takip", layout="wide")
    init_db()
    init_auth_state()

    is_manager = bool(st.session_state.manager_authenticated)

    st.title("POS Üretim ve Fire Takip Uygulaması")
    st.caption("İşçi çalışma süresi, POS, kombinasyon, adet ve fire bilgilerini kaydeder. — Sürüm 2.6 KOLAY ÇOKLU POS")
    if is_manager:
        st.success("Yönetici görünümü: Tüm kayıtlar, grafikler, raporlar ve silme işlemleri açık.")
    else:
        st.info("Çalışan görünümü: Yeni Kayıt ekranından OC ve POS seçip günlük üretimi girebilirsin.")

    worker_pages = ["Yeni Kayıt", "Üretime Devam Et"]
    manager_pages = [
        "Yeni Kayıt",
        "Üretime Devam Et",
        "Kayıtlar",
        "Grafikler",
        "Yönetici Paneli",
        "Gün Sonu Kontrolü",
        "Özet",
    ]

    with st.sidebar:
        st.header("Menü")
        st.info("Sürüm 2.6 KOLAY ÇOKLU POS")
        manager_login_panel()
        if is_manager:
            manager_overtime_sidebar_alert()
        st.divider()

        allowed_pages = manager_pages if is_manager else worker_pages
        page = st.radio("Sayfa", allowed_pages, index=0)

        uploaded_excel = None
        if is_manager:
            st.divider()
            sidebar_production_output_manager()
            st.divider()
            uploaded_excel = st.file_uploader(
                "Ayar Excel'ini değiştir",
                type=["xlsx"],
                help="Boş bırakırsan uygulama klasöründeki ayar_dosyasi.xlsx kullanılır.",
            )

    (reasons, combinations), source_info = load_config(uploaded_excel)
    if is_manager:
        st.sidebar.caption(source_info)

    if page == "Yeni Kayıt":
        if not combinations:
            st.info("Devam etmek için uygulama yöneticisinin ayar Excel'ini tanımlaması gerekir.")
            st.stop()
        yeni_kayit_page(reasons, combinations, is_manager)
    elif page == "Üretime Devam Et":
        if not combinations:
            st.info("Devam etmek için uygulama yöneticisinin ayar Excel'ini tanımlaması gerekir.")
            st.stop()
        uretime_devam_page(combinations)
    elif page == "Kayıtlar" and is_manager:
        kayitlar_page()
    elif page == "Grafikler" and is_manager:
        grafikler_page()
    elif page == "Yönetici Paneli" and is_manager:
        yonetici_paneli_page()
    elif page == "Gün Sonu Kontrolü" and is_manager:
        gun_sonu_kontrolu_page()
    elif page == "Özet" and is_manager:
        ozet_page()
    else:
        st.error("Bu bölümü görüntüleme yetkiniz yok.")

def yeni_kayit_page(reasons, combinations, is_manager: bool):
    combo_names = [c["ad"] for c in combinations]

    st.subheader("Yeni Üretim Kaydı")

    all_output_items = get_production_output_summary()
    if all_output_items.empty:
        st.warning(
            "Henüz yüklenmiş üretim çıktısı bulunmuyor. "
            "Yönetici sol menüden PDF veya Excel üretim çıktısı yüklemelidir."
        )
        return

    oc_table = (
        all_output_items[["oc_no", "project_name"]]
        .drop_duplicates()
        .sort_values(["oc_no", "project_name"])
        .reset_index(drop=True)
    )
    oc_labels = oc_table.apply(
        lambda row: f"OC {row['oc_no']} — {row['project_name']}",
        axis=1,
    ).tolist()
    label_to_oc = dict(zip(oc_labels, oc_table["oc_no"].astype(str)))

    top1, top2, top3 = st.columns(3)
    with top1:
        tarih = st.date_input("Tarih", value=date.today())
    with top2:
        operator_ismi = st.selectbox(
            "Operatör / İşçi ismi",
            WORKER_NAMES,
            format_func=lambda value: "İsim seçin" if value == "" else value,
        )
    with top3:
        selected_oc_label = st.selectbox(
            "OC seç",
            oc_labels,
            key="new_record_oc_select",
        )
        selected_oc = label_to_oc[selected_oc_label]

    selected_oc_items = get_production_output_summary(selected_oc)
    project_name = str(selected_oc_items["project_name"].iloc[0])

    st.info(f"**OC {selected_oc}** · {project_name}")

    st.divider()
    work1, work2, work3 = st.columns(3)
    with work1:
        calisma_tipi = st.selectbox(
            "Çalışma tipi",
            ["Tam zamanlı", "Yarı zamanlı", "Mesaili"],
        )

    default_hours = {
        "Tam zamanlı": 9.0,
        "Yarı zamanlı": 6.0,
        "Mesaili": 10.0,
    }[calisma_tipi]

    with work2:
        calisma_saati = st.number_input(
            "O gün toplam kaç saat çalıştı?",
            min_value=0.0,
            max_value=24.0,
            value=default_hours,
            step=0.5,
            key=f"work_hours_{calisma_tipi}",
        )
    with work3:
        neden = st.selectbox(
            "Yarı zamanlı çalışma nedeni",
            [""] + reasons,
            disabled=(calisma_tipi != "Yarı zamanlı"),
        )

    if calisma_tipi == "Mesaili":
        st.warning(
            "Bu kayıt mesaili olarak işaretlendi. "
            "Kaydedildiğinde yönetici ekranında mesai uyarısı görünecek."
        )

    multiple_pos = st.checkbox(
        "Bugün birden fazla POS ürettiniz mi?",
        value=False,
        help=(
            "Bir işçi aynı gün birden fazla POS ürettiyse bu seçeneği açıp "
            "POS'ları ve adetleri tek tabloda girebilir."
        ),
    )

    st.divider()

    if multiple_pos:
        st.subheader("Çoklu POS Günlük Üretim Girişi")

        all_pos_options = selected_oc_items["pos"].astype(str).tolist()
        selected_positions = st.multiselect(
            "Bugün üretilen POS'ları seç",
            all_pos_options,
            key=f"multi_pos_select_{selected_oc}",
        )

        if not selected_positions:
            st.info("Devam etmek için en az bir POS seç.")
            return

        chosen = selected_oc_items[
            selected_oc_items["pos"].astype(str).isin(selected_positions)
        ].copy()

        kombinasyon_adi = st.selectbox(
            "Bu POS'lar için kombinasyon",
            combo_names,
            key=f"multi_combo_{selected_oc}",
        )
        selected_combo = next(
            item for item in combinations
            if item["ad"] == kombinasyon_adi
        )
        operations = selected_combo["operasyonlar"]

        all_operations_done = st.checkbox(
            "Seçilen POS'larda bütün operasyonlar tamamlandı",
            value=True,
            key=f"multi_all_operations_done_{selected_oc}",
        )

        if all_operations_done:
            completed_operations = list(operations)
            st.success("Bütün operasyonlar tamamlandı olarak işaretlenecek.")
        else:
            completed_operations = st.multiselect(
                "Tamamlanan operasyonları seç",
                operations,
                default=[],
                key=f"multi_completed_ops_{selected_oc}",
            )

        st.markdown("### POS bazında günlük adet girişi")
        st.caption(
            "Her POS kartında yalnızca bugün üretilen adet ve varsa fire adedini gir. "
            "Kalanın tamamı üretildiyse 'Kalanın tamamını yaz' butonuna bas."
        )

        input_rows = []

        for _, pos_row in chosen.sort_values("pos").iterrows():
            item_id = int(pos_row["item_id"])
            pos_name = str(pos_row["pos"])
            requested_qty = int(pos_row["requested_qty"])
            produced_before = int(pos_row["produced_qty"])
            remaining_qty = int(pos_row["remaining_qty"])
            boy_mm = float(pos_row["boy_mm"])
            en_mm = float(pos_row["en_mm"])
            unit_area = float(pos_row["unit_area_mm2"])

            qty_key = f"multi_qty_{selected_oc}_{item_id}"
            fire_key = f"multi_fire_{selected_oc}_{item_id}"

            with st.container(border=True):
                title_col, status_col = st.columns([2, 1])
                with title_col:
                    st.markdown(f"## {pos_name}")
                    st.caption(
                        f"{boy_mm:,.2f} × {en_mm:,.2f} mm · "
                        f"Birim alan: {_format_mm2(unit_area)} mm²"
                    )
                with status_col:
                    if remaining_qty == 0:
                        st.success("Tamamlandı")
                    else:
                        st.warning(f"{remaining_qty} adet kaldı")

                m1, m2, m3 = st.columns(3)
                m1.metric("İstenen", f"{requested_qty} adet")
                m2.metric("Önceden üretilen", f"{produced_before} adet")
                m3.metric("Kalan", f"{remaining_qty} adet")

                input_col, fire_col, quick_col = st.columns([1.4, 1, 1.2])
                with input_col:
                    today_value = st.number_input(
                        "Bugün üretilen",
                        min_value=0,
                        max_value=max(remaining_qty, 0),
                        value=min(
                            int(st.session_state.get(qty_key, 0)),
                            max(remaining_qty, 0),
                        ),
                        step=1,
                        key=qty_key,
                    )
                with fire_col:
                    fire_value = st.number_input(
                        "Fire adedi",
                        min_value=0,
                        value=int(st.session_state.get(fire_key, 0)),
                        step=1,
                        key=fire_key,
                    )
                with quick_col:
                    st.write("")
                    st.write("")
                    st.button(
                        "Kalanın tamamını yaz",
                        disabled=(remaining_qty == 0),
                        use_container_width=True,
                        key=f"fill_remaining_{selected_oc}_{item_id}",
                        on_click=set_session_value,
                        args=(qty_key, remaining_qty),
                    )

                st.caption(
                    f"Bugün üretilecek alan: "
                    f"{_format_mm2(int(today_value) * unit_area)} mm²"
                )

                input_rows.append(
                    {
                        "item_id": item_id,
                        "POS": pos_name,
                        "İstenen Adet": requested_qty,
                        "Önceden Üretilen": produced_before,
                        "Kalan": remaining_qty,
                        "Boy (mm)": boy_mm,
                        "En (mm)": en_mm,
                        "Birim Alan (mm²)": unit_area,
                        "Bugün Üretilen": int(today_value),
                        "Fire Adedi": int(fire_value),
                    }
                )

        edited = pd.DataFrame(input_rows)

        today_qty = pd.to_numeric(
            edited["Bugün Üretilen"],
            errors="coerce",
        ).fillna(0).astype(int)
        fire_qty = pd.to_numeric(
            edited["Fire Adedi"],
            errors="coerce",
        ).fillna(0).astype(int)

        total_today = int(today_qty.sum())
        total_fire = int(fire_qty.sum())
        total_area = float(
            (
                today_qty
                * pd.to_numeric(
                    edited["Birim Alan (mm²)"],
                    errors="coerce",
                ).fillna(0)
            ).sum()
        )

        summary1, summary2, summary3 = st.columns(3)
        summary1.metric("Bugün üretilen toplam", f"{total_today} adet")
        summary2.metric("Toplam fire", f"{total_fire} adet")
        summary3.metric("Bugün üretilen alan", f"{_format_mm2(total_area)} mm²")

        fire_operation = ""
        fire_note = ""
        if total_fire > 0:
            fire1, fire2 = st.columns(2)
            with fire1:
                fire_operation = st.selectbox(
                    "Fire hangi operasyonda oluştu?",
                    [""] + operations,
                    key=f"multi_fire_operation_{selected_oc}",
                )
            with fire2:
                fire_note = st.text_input(
                    "Fire açıklaması",
                    key=f"multi_fire_note_{selected_oc}",
                )

        general_note = st.text_area(
            "Genel not",
            key=f"multi_note_{selected_oc}",
        )

        if st.button(
            "Çoklu POS Kayıtlarını Kaydet",
            type="primary",
            key=f"save_multi_pos_{selected_oc}",
        ):
            errors = []

            if not operator_ismi:
                errors.append("Operatör / işçi seçmelisin.")
            if calisma_tipi == "Yarı zamanlı" and not neden:
                errors.append("Yarı zamanlı çalışma için neden seçmelisin.")
            if calisma_tipi == "Mesaili" and float(calisma_saati) <= 9.0:
                errors.append(
                    "Mesaili kayıt için çalışma saati 9 saatten fazla olmalıdır."
                )
            if total_today <= 0:
                errors.append("En az bir POS için bugün üretilen adet girmelisin.")

            over_rows = edited[
                today_qty > edited["Kalan"].astype(int)
            ]
            if not over_rows.empty:
                errors.append(
                    "Bugün üretilen adet kalan miktarı aşamaz: "
                    + ", ".join(over_rows["POS"].astype(str).tolist())
                )

            if total_fire > 0 and not fire_operation:
                errors.append("Fire varsa fire operasyonunu seçmelisin.")

            if errors:
                for error in errors:
                    st.error(error)
                st.stop()

            active_mask = (today_qty + fire_qty) > 0
            active_rows = edited.loc[active_mask].copy()
            active_today = today_qty.loc[active_mask]
            active_fire = fire_qty.loc[active_mask]
            workloads = active_today + active_fire
            total_workload = int(workloads.sum())

            allocated_hours = []
            remaining_hours = float(calisma_saati)
            for index, workload in enumerate(workloads.tolist()):
                if index == len(workloads) - 1:
                    allocated = remaining_hours
                else:
                    allocated = (
                        float(calisma_saati)
                        * float(workload)
                        / float(total_workload)
                    )
                    allocated = round(allocated, 4)
                    remaining_hours -= allocated
                allocated_hours.append(max(allocated, 0.0))

            session_ids = []
            output_log_rows = []

            for row_position, (_, row) in enumerate(active_rows.iterrows()):
                produced = int(active_today.loc[row.name])
                fire = int(active_fire.loc[row.name])
                workload = produced + fire

                operation_rows = build_quick_operation_rows(
                    produced_qty=produced,
                    fire_qty=fire,
                    operations=operations,
                    completed_operations=completed_operations,
                    fire_operation=fire_operation,
                    fire_note=fire_note,
                )

                note_parts = [
                    "Çoklu POS günlük kaydı",
                    f"Günlük toplam çalışma: {float(calisma_saati):.1f} saat",
                ]
                if general_note.strip():
                    note_parts.append(general_note.strip())

                session_data = {
                    "tarih": str(tarih),
                    "operator_ismi": operator_ismi.strip(),
                    "proje": str(selected_oc),
                    "pos": str(row["POS"]),
                    "kombinasyon_adi": kombinasyon_adi,
                    "calisma_tipi": calisma_tipi,
                    "calisma_saati": float(allocated_hours[row_position]),
                    "neden": neden,
                    "siparis_adedi": int(row["İstenen Adet"]),
                    "saglam_tamamlanan": produced,
                    "fire_adedi": fire,
                    "uretim_yuku": workload,
                    "notlar": " | ".join(note_parts),
                }
                session_ids.append(
                    insert_session(session_data, operation_rows)
                )
                output_log_rows.append(
                    {
                        "item_id": int(row["item_id"]),
                        "Bugün Üretilen": produced,
                    }
                )

            output_frame = pd.DataFrame(output_log_rows)
            save_production_output_entries(
                output_frame,
                str(tarih),
                operator_ismi.strip(),
                general_note,
            )

            st.success(
                f"{len(session_ids)} POS kaydı eklendi. "
                f"Toplam {total_today} sağlam adet ve {total_fire} fire kaydedildi."
            )
        return

    # --------------------------------------------------------------
    # TEK POS DETAYLI GİRİŞ
    # --------------------------------------------------------------
    st.subheader("Tek POS Detaylı Üretim Girişi")

    pos_options = selected_oc_items["pos"].astype(str).tolist()
    pos = st.selectbox(
        "POS seç",
        pos_options,
        key=f"new_record_pos_{selected_oc}",
    )
    selected_item = selected_oc_items[
        selected_oc_items["pos"].astype(str) == str(pos)
    ].iloc[0]

    siparis_adedi = int(selected_item["requested_qty"])
    boy_mm = float(selected_item["boy_mm"])
    en_mm = float(selected_item["en_mm"])
    unit_area_mm2 = float(selected_item["unit_area_mm2"])
    requested_area_mm2 = float(selected_item["requested_area_mm2"])
    output_produced_qty = int(selected_item["produced_qty"])
    output_remaining_qty = int(selected_item["remaining_qty"])

    st.info(
        f"**OC {selected_oc}** · {project_name} · **{pos}** · "
        f"İstenen: **{siparis_adedi} adet**"
    )

    detail1, detail2, detail3, detail4 = st.columns(4)
    detail1.metric("Boy", f"{_format_mm2(boy_mm)} mm")
    detail2.metric("En", f"{_format_mm2(en_mm)} mm")
    detail3.metric("Birim alan", f"{_format_mm2(unit_area_mm2)} mm²")
    detail4.metric(
        "İstenen toplam alan",
        f"{_format_mm2(requested_area_mm2)} mm²",
    )

    if is_manager:
        manager1, manager2, manager3 = st.columns(3)
        manager1.metric(
            "Önceden üretilen",
            f"{output_produced_qty} adet",
        )
        manager2.metric(
            "Kalan",
            f"{output_remaining_qty} adet",
        )
        manager3.metric(
            "Kalan alan",
            f"{_format_mm2(selected_item['remaining_area_mm2'])} mm²",
        )

    combo1, combo2 = st.columns(2)
    with combo1:
        kombinasyon_adi = st.selectbox(
            "Kombinasyon",
            combo_names,
        )
    selected_combo = next(
        item for item in combinations
        if item["ad"] == kombinasyon_adi
    )
    operations = selected_combo["operasyonlar"]
    with combo2:
        st.write("Seçilen operasyonlar")
        st.write(" → ".join(operations))

    st.divider()

    values1, values2, values3 = st.columns(3)
    with values1:
        st.number_input(
            "İstenen üretim adedi",
            min_value=0,
            value=siparis_adedi,
            step=1,
            disabled=True,
            key=f"requested_qty_{selected_oc}_{pos}",
        )
    with values2:
        saglam_tamamlanan = st.number_input(
            "Bugün üretilen sağlam adet",
            min_value=0,
            max_value=max(output_remaining_qty, 0),
            value=0,
            step=1,
        )
    with values3:
        fire_adedi = st.number_input(
            "Fire adedi",
            min_value=0,
            value=0,
            step=1,
        )

    uretim_yuku = int(saglam_tamamlanan + fire_adedi)
    kalan = max(output_remaining_qty - saglam_tamamlanan, 0)
    produced_today_area = unit_area_mm2 * int(saglam_tamamlanan)

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("Bugünkü üretim yükü", f"{uretim_yuku} adet")
    metric2.metric("Kalan ihtiyaç", f"{kalan} adet")
    metric3.metric(
        "Bugün üretilen alan",
        f"{_format_mm2(produced_today_area)} mm²",
    )

    notlar = st.text_area("Genel not")

    st.divider()
    st.subheader("Adet Bazlı Operasyon Takibi")
    st.caption(
        "Bugün üretilen sağlam adet ve fire toplamına göre ürünler açılır."
    )

    operation_rows = []
    fire_options = [""] + operations
    tracked_unit_count = max(uretim_yuku, 1)

    st.checkbox(
        "Tüm ürünlerde tüm operasyonları seç",
        key="select_all_all_units",
        on_change=set_all_operations_for_all_units,
        args=(tracked_unit_count, len(operations)),
    )

    for unit_no in range(1, tracked_unit_count + 1):
        with st.expander(
            f"Ürün / Adet {unit_no}",
            expanded=(unit_no <= 2),
        ):
            st.checkbox(
                "Bu ürün için tüm operasyonları seç",
                key=f"select_all_{unit_no}",
                on_change=set_all_operations_for_unit,
                args=(unit_no, len(operations)),
            )

            top_columns = st.columns([1, 2, 3])
            with top_columns[0]:
                fire_var = st.checkbox(
                    "Bu adette fire var",
                    key=f"fire_{unit_no}",
                )
            with top_columns[1]:
                fire_operasyonu = st.selectbox(
                    "Fire hangi operasyonda oldu?",
                    fire_options,
                    key=f"fireop_{unit_no}",
                    disabled=not fire_var,
                )
            with top_columns[2]:
                fire_notu = st.text_input(
                    "Fire notu",
                    key=f"firenote_{unit_no}",
                    disabled=not fire_var,
                )

            columns = st.columns(min(max(len(operations), 1), 5))
            for operation_index, operation_name in enumerate(
                operations,
                start=1,
            ):
                with columns[(operation_index - 1) % len(columns)]:
                    done = st.checkbox(
                        operation_name,
                        key=f"op_{unit_no}_{operation_index}",
                    )

                operation_rows.append(
                    {
                        "urun_no": unit_no,
                        "operasyon_sirasi": operation_index,
                        "operasyon_adi": operation_name,
                        "yapildi": done,
                        "fire_var": fire_var,
                        "fire_operasyonu": (
                            fire_operasyonu if fire_var else ""
                        ),
                        "fire_notu": fire_notu if fire_var else "",
                    }
                )

    st.divider()

    if st.button("Kaydı Veritabanına Ekle", type="primary"):
        errors = []

        if not operator_ismi:
            errors.append("Operatör / işçi seçmelisin.")
        if calisma_tipi == "Yarı zamanlı" and not neden:
            errors.append("Yarı zamanlı çalışma için neden seçmelisin.")
        if calisma_tipi == "Mesaili" and float(calisma_saati) <= 9.0:
            errors.append(
                "Mesaili kayıt için çalışma saati 9 saatten fazla olmalıdır."
            )
        if saglam_tamamlanan <= 0:
            errors.append("Bugün üretilen sağlam adedi girmelisin.")
        if saglam_tamamlanan > output_remaining_qty:
            errors.append(
                "Bugün üretilen adet kalan miktardan büyük olamaz."
            )

        fire_marked_units = len(
            {
                row["urun_no"]
                for row in operation_rows
                if row["fire_var"]
            }
        )
        if fire_marked_units != fire_adedi:
            errors.append(
                f"Fire adedi {fire_adedi} girildi ancak "
                f"{fire_marked_units} ürün fireli işaretlendi."
            )

        if errors:
            for error in errors:
                st.error(error)
            st.stop()

        session_data = {
            "tarih": str(tarih),
            "operator_ismi": operator_ismi.strip(),
            "proje": str(selected_oc),
            "pos": pos,
            "kombinasyon_adi": kombinasyon_adi,
            "calisma_tipi": calisma_tipi,
            "calisma_saati": float(calisma_saati),
            "neden": neden,
            "siparis_adedi": siparis_adedi,
            "saglam_tamamlanan": int(saglam_tamamlanan),
            "fire_adedi": int(fire_adedi),
            "uretim_yuku": int(uretim_yuku),
            "notlar": notlar,
        }

        session_id = insert_session(session_data, operation_rows)

        output_frame = pd.DataFrame(
            [
                {
                    "item_id": int(selected_item["item_id"]),
                    "Bugün Üretilen": int(saglam_tamamlanan),
                }
            ]
        )
        save_production_output_entries(
            output_frame,
            str(tarih),
            operator_ismi.strip(),
            notlar,
        )

        st.success(
            f"OC {selected_oc} / {pos} kaydı eklendi. "
            f"Kayıt no: {session_id}"
        )





def uretim_ciktisi_page(is_manager: bool):
    st.subheader("Ana Ekran — Üretim Çıktısı")
    st.caption(
        "Çalışan OC No seçer. O OC'ye ait POS, istenen adet, Boy, En ve alan bilgileri "
        "otomatik gelir. Çalışan yalnızca bugün ürettiği adetleri yazar."
    )

    success_message = st.session_state.pop("production_output_success", None)
    if success_message:
        st.success(success_message)

    if is_manager:
        with st.expander("Yönetici: Üretim çıktısı dosyalarını içe aktar", expanded=False):
            uploaded_outputs = st.file_uploader(
                "PDF veya Excel üretim çıktıları",
                type=["pdf", "xlsx", "xls"],
                accept_multiple_files=True,
                key="production_output_uploads",
                help="Aynı anda birden fazla dosya seçebilirsin.",
            )

            parsed_files = []
            if uploaded_outputs:
                total_rows = 0
                total_requested = 0
                total_projects = set()

                for uploaded_output in uploaded_outputs:
                    try:
                        parsed = parse_production_output_file(
                            uploaded_output.getvalue(),
                            uploaded_output.name,
                        )
                        parsed_files.append((uploaded_output.name, parsed))
                        total_rows += len(parsed)
                        total_requested += int(parsed["requested_qty"].sum())
                        total_projects.update(
                            zip(
                                parsed["project_name"].astype(str),
                                parsed["oc_no"].astype(str),
                            )
                        )

                        project_names = (
                            parsed[["project_name", "oc_no"]]
                            .drop_duplicates()
                            .apply(
                                lambda row: f"{row['project_name']} — OC {row['oc_no']}",
                                axis=1,
                            )
                            .tolist()
                        )

                        with st.expander(
                            f"{uploaded_output.name} · {', '.join(project_names)}",
                            expanded=False,
                        ):
                            preview = parsed.rename(columns={
                                "project_name": "Proje",
                                "oc_no": "OC",
                                "pos": "POS",
                                "requested_qty": "İstenen Adet",
                                "boy_mm": "Boy (mm)",
                                "en_mm": "En (mm)",
                                "unit_area_mm2": "Birim Alan (mm²)",
                            })
                            st.dataframe(
                                preview,
                                use_container_width=True,
                                hide_index=True,
                            )
                    except Exception as exc:
                        st.error(f"{uploaded_output.name} okunamadı: {exc}")

                if parsed_files:
                    st.success(
                        f"{len(parsed_files)} dosyada {len(total_projects)} proje/OC ve "
                        f"{total_rows} POS satırı okundu. "
                        f"Toplam istenen adet: {total_requested}."
                    )
                    if st.button(
                        "Tüm Dosyaları Sisteme Aktar",
                        type="primary",
                        key="import_all_production_outputs",
                    ):
                        imported_count = 0
                        for source_name, parsed in parsed_files:
                            imported_count += import_production_output_items(
                                parsed,
                                source_name,
                            )
                        st.session_state.production_output_success = (
                            f"{len(parsed_files)} dosyadan {imported_count} POS satırı "
                            "sisteme aktarıldı veya güncellendi."
                        )
                        st.rerun()
        st.divider()
    else:
        st.info("Dosya yükleme işlemi yalnızca yönetici görünümünde açıktır.")

    all_items = get_production_output_summary()
    if all_items.empty:
        st.info("Henüz sisteme aktarılmış üretim çıktısı yok.")
        return

    oc_choices = (
        all_items[["oc_no", "project_name"]]
        .drop_duplicates()
        .sort_values(["oc_no", "project_name"])
        .reset_index(drop=True)
    )
    oc_labels = oc_choices.apply(
        lambda row: f"OC {row['oc_no']} — {row['project_name']}",
        axis=1,
    ).tolist()
    label_to_oc = dict(zip(oc_labels, oc_choices["oc_no"].astype(str)))

    selected_oc_label = st.selectbox(
        "OC No seç",
        oc_labels,
        key="production_output_oc",
    )
    selected_oc = label_to_oc[selected_oc_label]
    items = get_production_output_summary(selected_oc)

    selected_project_name = str(items["project_name"].iloc[0])
    st.info(f"OC: **{selected_oc}** · Proje: **{selected_project_name}**")

    total_requested = int(items["requested_qty"].sum())
    total_produced = int(items["produced_qty"].sum())
    total_remaining = int(items["remaining_qty"].sum())
    requested_area = float(items["requested_area_mm2"].sum())
    produced_area = float(items["produced_area_mm2"].sum())
    remaining_area = float(items["remaining_area_mm2"].sum())

    m1, m2, m3 = st.columns(3)
    m1.metric("İstenen toplam adet", total_requested)
    m2.metric("Üretilen toplam adet", total_produced)
    m3.metric("Kalan toplam adet", total_remaining)

    a1, a2, a3 = st.columns(3)
    a1.metric("İstenen alan", f"{_format_mm2(requested_area)} mm²")
    a2.metric("Üretilen alan", f"{_format_mm2(produced_area)} mm²")
    a3.metric("Kalan alan", f"{_format_mm2(remaining_area)} mm²")

    st.divider()
    c1, c2, c3 = st.columns(3)
    with c1:
        production_date = st.date_input(
            "Üretim tarihi",
            value=date.today(),
            key=f"output_date_{selected_oc}",
        )
    with c2:
        operator_name = st.selectbox(
            "Operatör / İşçi",
            WORKER_NAMES,
            format_func=lambda value: "İsim seçin" if value == "" else value,
            key=f"output_operator_{selected_oc}",
        )
    with c3:
        only_remaining = st.checkbox(
            "Sadece kalan POS'ları göster",
            value=True,
            key=f"output_remaining_{selected_oc}",
        )

    note = st.text_input("Not", key=f"output_note_{selected_oc}")

    editable = items.copy()
    if only_remaining:
        editable = editable[editable["remaining_qty"] > 0].copy()

    if editable.empty:
        st.success("Bu proje/OC içindeki bütün POS adetleri tamamlandı.")
    else:
        editor = pd.DataFrame({
            "item_id": editable["item_id"].astype(int),
            "OC": editable["oc_no"].astype(str),
            "POS": editable["pos"].astype(str),
            "İstenen Adet": editable["requested_qty"].astype(int),
            "Boy (mm)": editable["boy_mm"].astype(float),
            "En (mm)": editable["en_mm"].astype(float),
            "Birim Alan (mm²)": editable["unit_area_mm2"].round(2),
            "Önceden Üretilen": editable["produced_qty"].astype(int),
            "Kalan": editable["remaining_qty"].astype(int),
            "Bugün Üretilen": 0,
        })

        edited = st.data_editor(
            editor,
            use_container_width=True,
            hide_index=True,
            disabled=[
                "OC", "POS", "İstenen Adet", "Boy (mm)", "En (mm)",
                "Birim Alan (mm²)", "Önceden Üretilen", "Kalan",
            ],
            column_config={
                "item_id": None,
                "Bugün Üretilen": st.column_config.NumberColumn(
                    "Bugün Üretilen",
                    min_value=0,
                    step=1,
                    required=True,
                ),
                "Birim Alan (mm²)": st.column_config.NumberColumn(format="%.2f"),
                "Boy (mm)": st.column_config.NumberColumn(format="%.2f"),
                "En (mm)": st.column_config.NumberColumn(format="%.2f"),
            },
            key=f"output_editor_{selected_oc}",
        )

        entered = pd.to_numeric(
            edited["Bugün Üretilen"],
            errors="coerce",
        ).fillna(0).astype(int)
        entered_area = float(
            (entered * edited["Birim Alan (mm²)"].astype(float)).sum()
        )

        e1, e2 = st.columns(2)
        e1.metric("Bugün girilen adet", int(entered.sum()))
        e2.metric(
            "Bugün üretilen alan",
            f"{_format_mm2(entered_area)} mm²",
        )

        if st.button(
            "Üretim Miktarlarını Kaydet",
            type="primary",
            key=f"save_output_{selected_oc}",
        ):
            errors = []
            if not operator_name:
                errors.append("Operatör / işçi seçmelisin.")

            over_rows = edited[entered > edited["Kalan"].astype(int)]
            if not over_rows.empty:
                errors.append(
                    "Girilen üretim adedi kalan miktarı aşamaz: "
                    + ", ".join(over_rows["POS"].astype(str).tolist())
                )

            if int(entered.sum()) <= 0:
                errors.append("En az bir POS için üretilen adet girmelisin.")

            if errors:
                for error in errors:
                    st.error(error)
            else:
                count = save_production_output_entries(
                    edited,
                    str(production_date),
                    operator_name,
                    note,
                )
                st.session_state.production_output_success = (
                    f"{selected_project_name} / OC {selected_oc}: "
                    f"{count} POS için üretim miktarı kaydedildi."
                )
                st.rerun()

    if is_manager:
        st.divider()
        st.subheader("Proje / OC Üretim Geçmişi")
        history = get_production_output_history(selected_oc)
        if history.empty:
            st.info("Bu proje/OC için henüz üretim girişi yok.")
        else:
            st.dataframe(history, use_container_width=True, hide_index=True)
            csv = history.to_csv(index=False).encode("utf-8-sig")
            safe_project = re.sub(
                r"[^A-Za-z0-9_-]+",
                "_",
                selected_project_name,
            ).strip("_")
            st.download_button(
                "Üretim Geçmişini CSV İndir",
                data=csv,
                file_name=(
                    f"{safe_project or 'proje'}_OC_{selected_oc}_"
                    "uretim_gecmisi.csv"
                ),
                mime="text/csv",
            )



def uretime_devam_page(combinations):
    st.subheader("Üretime Devam Et / Kaydı Güncelle")
    st.caption(
        "Tamamlanmamış bir POS'u aç, daha önce yapılan adımları koru, "
        "kalan operasyonları işaretle ve kaydı güncelle."
    )

    success_message = st.session_state.pop("progress_update_success", None)
    if success_message:
        st.success(success_message)

    conn = sqlite3.connect(DB_PATH)
    sessions = pd.read_sql_query(
        "SELECT * FROM work_sessions ORDER BY id DESC",
        conn,
    )
    conn.close()

    if sessions.empty:
        st.info("Devam ettirilecek kayıt yok. Önce yeni üretim kaydı eklemelisin.")
        return

    sessions["durum"] = sessions.apply(
        lambda row: production_status(
            int(row["siparis_adedi"]),
            int(row["saglam_tamamlanan"]),
        ),
        axis=1,
    )
    sessions["kalan"] = (
        sessions["siparis_adedi"] - sessions["saglam_tamamlanan"]
    ).clip(lower=0)

    only_incomplete = st.checkbox(
        "Sadece tamamlanmamış kayıtları göster",
        value=True,
        key="continue_only_incomplete",
    )
    visible_sessions = sessions.copy()
    if only_incomplete:
        visible_sessions = visible_sessions[
            visible_sessions["saglam_tamamlanan"]
            < visible_sessions["siparis_adedi"]
        ].copy()

    if visible_sessions.empty:
        st.success("Tamamlanmamış kayıt bulunmuyor.")
        return

    labels = visible_sessions.apply(
        lambda row: (
            f"#{int(row['id'])} | {row['pos']} | {row['operator_ismi']} | "
            f"{int(row['saglam_tamamlanan'])}/{int(row['siparis_adedi'])} | "
            f"{row['durum']}"
        ),
        axis=1,
    ).tolist()
    selected_label = st.selectbox(
        "Devam etmek istediğin kayıt",
        labels,
        key="continue_record_select",
    )
    selected_index = labels.index(selected_label)
    selected_session = visible_sessions.iloc[selected_index]
    selected_id = int(selected_session["id"])

    conn = sqlite3.connect(DB_PATH)
    details = pd.read_sql_query(
        """
        SELECT
            urun_no,
            operasyon_sirasi,
            operasyon_adi,
            yapildi,
            fire_var,
            fire_operasyonu,
            fire_notu
        FROM operation_entries
        WHERE session_id = ?
        ORDER BY urun_no, operasyon_sirasi
        """,
        conn,
        params=(selected_id,),
    )
    conn.close()

    status = production_status(
        int(selected_session["siparis_adedi"]),
        int(selected_session["saglam_tamamlanan"]),
    )
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("POS", selected_session["pos"])
    m2.metric("Sipariş", int(selected_session["siparis_adedi"]))
    m3.metric("Tamamlanan", int(selected_session["saglam_tamamlanan"]))
    m4.metric("Kalan", int(selected_session["kalan"]))
    m5.metric("Durum", status)

    st.caption(
        f"OC: {selected_session['proje'] or '-'} | "
        f"Kombinasyon: {selected_session['kombinasyon_adi']} | "
        f"Toplam çalışma: {float(selected_session['calisma_saati']):.1f} saat"
    )

    combination = next(
        (
            combo
            for combo in combinations
            if combo["ad"] == selected_session["kombinasyon_adi"]
        ),
        None,
    )
    if not details.empty:
        operations = (
            details[["operasyon_sirasi", "operasyon_adi"]]
            .drop_duplicates()
            .sort_values("operasyon_sirasi")["operasyon_adi"]
            .astype(str)
            .tolist()
        )
    elif combination:
        operations = [str(op) for op in combination["operasyonlar"]]
    else:
        st.error("Bu kayıt için operasyon listesi bulunamadı.")
        return

    previous_map = {
        (int(row["urun_no"]), int(row["operasyon_sirasi"])): bool(row["yapildi"])
        for _, row in details.iterrows()
    }
    fire_map = {}
    if not details.empty:
        for urun_no, group in details.groupby("urun_no"):
            first = group.iloc[0]
            fire_map[int(urun_no)] = {
                "fire_var": bool(first["fire_var"]),
                "fire_operasyonu": ""
                if pd.isna(first["fire_operasyonu"])
                else str(first["fire_operasyonu"]),
                "fire_notu": ""
                if pd.isna(first["fire_notu"])
                else str(first["fire_notu"]),
            }

    st.divider()
    st.subheader("Güncelleme Bilgileri")
    u1, u2, u3, u4 = st.columns(4)
    with u1:
        update_date = st.date_input(
            "Güncelleme tarihi",
            value=date.today(),
            key=f"edit_date_{selected_id}",
        )
    with u2:
        updated_by = st.text_input(
            "Güncelleyen işçi / yönetici",
            key=f"edit_by_{selected_id}",
        )
    with u3:
        added_hours = st.number_input(
            "Bu güncellemede ek çalışma saati",
            min_value=0.0,
            max_value=24.0,
            value=0.0,
            step=0.5,
            key=f"edit_hours_{selected_id}",
        )
    with u4:
        new_completed = st.number_input(
            "Yeni sağlam tamamlanan adet",
            min_value=0,
            max_value=int(selected_session["siparis_adedi"]),
            value=int(selected_session["saglam_tamamlanan"]),
            step=1,
            key=f"edit_completed_{selected_id}",
        )

    new_fire = st.number_input(
        "Güncel toplam fire adedi",
        min_value=0,
        value=int(selected_session["fire_adedi"]),
        step=1,
        key=f"edit_fire_count_{selected_id}",
        help="Yeni fire oluştuysa toplam fire sayısını artır. Üretim yükü sipariş + fire olarak güncellenir.",
    )
    new_load = int(selected_session["siparis_adedi"]) + int(new_fire)
    st.info(
        f"Güncel üretim yükü: {int(selected_session['siparis_adedi'])} sipariş + "
        f"{int(new_fire)} fire = {new_load} adet"
    )
    update_note = st.text_area(
        "Güncelleme açıklaması",
        key=f"edit_note_{selected_id}",
        placeholder="Örnek: Kalan 2 ürünün boya ve paketleme operasyonları tamamlandı.",
    )

    st.divider()
    st.subheader("Operasyonları Güncelle")
    st.caption(
        "Daha önce yapılmış operasyonlar işaretli gelir. Yeni tamamlanan adımları işaretle."
    )

    all_key = f"edit_select_all_all_{selected_id}"
    if all_key not in st.session_state:
        st.session_state[all_key] = False
    st.checkbox(
        "Tüm ürünlerde tüm operasyonları seç",
        key=all_key,
        on_change=set_all_edit_operations_for_all_units,
        args=(selected_id, new_load, len(operations)),
    )

    operation_rows = []
    fire_options = [""] + operations

    for urun_no in range(1, max(new_load, 1) + 1):
        unit_existing = fire_map.get(
            urun_no,
            {"fire_var": False, "fire_operasyonu": "", "fire_notu": ""},
        )
        with st.expander(
            f"Ürün / Adet {urun_no}",
            expanded=(urun_no <= 2),
        ):
            unit_all_key = f"edit_select_all_{selected_id}_{urun_no}"
            if unit_all_key not in st.session_state:
                existing_values = [
                    previous_map.get((urun_no, idx), False)
                    for idx in range(1, len(operations) + 1)
                ]
                st.session_state[unit_all_key] = bool(existing_values) and all(
                    existing_values
                )
            st.checkbox(
                "Bu ürün için tüm operasyonları seç",
                key=unit_all_key,
                on_change=set_all_edit_operations_for_unit,
                args=(selected_id, urun_no, len(operations)),
            )

            fire_key = f"edit_fire_{selected_id}_{urun_no}"
            fireop_key = f"edit_fireop_{selected_id}_{urun_no}"
            firenote_key = f"edit_firenote_{selected_id}_{urun_no}"
            if fire_key not in st.session_state:
                st.session_state[fire_key] = bool(unit_existing["fire_var"])
            if fireop_key not in st.session_state:
                current_fire_op = unit_existing["fire_operasyonu"]
                st.session_state[fireop_key] = (
                    current_fire_op if current_fire_op in fire_options else ""
                )
            if firenote_key not in st.session_state:
                st.session_state[firenote_key] = unit_existing["fire_notu"]

            top_cols = st.columns([1, 2, 3])
            with top_cols[0]:
                fire_var = st.checkbox(
                    "Bu adette fire var",
                    key=fire_key,
                )
            with top_cols[1]:
                fire_operasyonu = st.selectbox(
                    "Fire hangi operasyonda oldu?",
                    fire_options,
                    key=fireop_key,
                    disabled=not fire_var,
                )
            with top_cols[2]:
                fire_notu = st.text_input(
                    "Fire notu",
                    key=firenote_key,
                    disabled=not fire_var,
                )

            cols = st.columns(min(max(len(operations), 1), 5))
            for idx, op in enumerate(operations, start=1):
                op_key = f"edit_op_{selected_id}_{urun_no}_{idx}"
                if op_key not in st.session_state:
                    st.session_state[op_key] = previous_map.get(
                        (urun_no, idx),
                        False,
                    )
                with cols[(idx - 1) % len(cols)]:
                    done = st.checkbox(op, key=op_key)

                operation_rows.append(
                    {
                        "urun_no": urun_no,
                        "operasyon_sirasi": idx,
                        "operasyon_adi": op,
                        "yapildi": done,
                        "fire_var": fire_var,
                        "fire_operasyonu": fire_operasyonu if fire_var else "",
                        "fire_notu": fire_notu if fire_var else "",
                    }
                )

    st.divider()
    new_status = production_status(
        int(selected_session["siparis_adedi"]),
        int(new_completed),
    )
    p1, p2, p3 = st.columns(3)
    p1.metric("Yeni durum", new_status)
    p2.metric(
        "Yeni kalan",
        max(int(selected_session["siparis_adedi"]) - int(new_completed), 0),
    )
    p3.metric(
        "Yeni toplam çalışma",
        f"{float(selected_session['calisma_saati']) + float(added_hours):.1f} saat",
    )

    if st.button(
        "Kaydı Güncelle ve Üretime Devam Et",
        type="primary",
        key=f"update_progress_button_{selected_id}",
    ):
        errors = []
        if not updated_by.strip():
            errors.append("Güncelleyen işçi / yönetici adı boş olamaz.")
        if int(new_completed) < int(selected_session["saglam_tamamlanan"]):
            errors.append(
                "Tamamlanan adet önceki değerden küçük olamaz. Hatalı kayıt düzeltmesi gerekiyorsa kaydı silip yeniden gir."
            )
        fire_marked_units = len(
            {
                int(row["urun_no"])
                for row in operation_rows
                if bool(row["fire_var"])
            }
        )
        if fire_marked_units != int(new_fire):
            errors.append(
                f"Toplam fire {int(new_fire)} girildi ancak ürünlerde "
                f"{fire_marked_units} adet fire işaretlendi."
            )

        if errors:
            for error in errors:
                st.error(error)
        else:
            update_session_progress(
                session_id=selected_id,
                new_completed=int(new_completed),
                new_fire=int(new_fire),
                added_hours=float(added_hours),
                updated_by=updated_by,
                update_date=str(update_date),
                update_note=update_note,
                operation_rows=operation_rows,
            )
            clear_edit_widget_state(selected_id)
            st.session_state["progress_update_success"] = (
                f"#{selected_id} numaralı {selected_session['pos']} kaydı güncellendi. "
                f"Yeni durum: {new_status}."
            )
            st.rerun()

    st.divider()
    st.subheader("Bu Kaydın Güncelleme Geçmişi")
    history = get_update_history(selected_id)
    if history.empty:
        st.info("Bu kayıt için henüz güncelleme geçmişi yok.")
    else:
        st.dataframe(history, use_container_width=True)


def render_operation_progress(details: pd.DataFrame, selected_session: pd.Series, selected_id: int):
    """Seçilen kaydın operasyon bazlı yapılan/kalan durumunu güvenli biçimde gösterir."""
    st.divider()
    st.subheader("Kombinasyon İlerleme Özeti")

    if details.empty:
        st.info("Bu kayda bağlı operasyon detayı bulunamadığı için grafik oluşturulamadı.")
        return

    toplam_urun = int(selected_session.get("uretim_yuku", 0) or 0)
    if toplam_urun <= 0:
        st.info("Toplam takip edilen adet 0 olduğu için grafik oluşturulamadı.")
        return

    combo_summary = (
        details.groupby("operasyon_adi", as_index=False)["yapildi"]
        .sum()
        .rename(columns={"yapildi": "yapilan_adet"})
    )

    if combo_summary.empty:
        st.info("Grafik için operasyon bilgisi bulunamadı.")
        return

    combo_summary["yapilan_adet"] = pd.to_numeric(
        combo_summary["yapilan_adet"], errors="coerce"
    ).fillna(0).astype(int)
    combo_summary["toplam_urun"] = toplam_urun
    combo_summary["kalan_adet"] = (
        combo_summary["toplam_urun"] - combo_summary["yapilan_adet"]
    ).clip(lower=0)
    combo_summary["tamamlanma_yuzdesi"] = (
        combo_summary["yapilan_adet"] / toplam_urun * 100
    ).round(1)

    st.caption(
        f"Seçilen kayıt: {selected_session['pos']} / "
        f"{selected_session['kombinasyon_adi']} - "
        f"Toplam takip edilen adet: {toplam_urun}"
    )
    st.dataframe(combo_summary, use_container_width=True)

    operasyonlar = combo_summary["operasyon_adi"].dropna().astype(str).tolist()
    if not operasyonlar:
        st.info("Seçilebilecek operasyon bulunamadı.")
        return

    grafik_operasyon = st.selectbox(
        "Daire grafiğinde görmek istediğin operasyon",
        operasyonlar,
        key=f"pie_select_{selected_id}",
    )

    selected_rows = combo_summary[
        combo_summary["operasyon_adi"].astype(str) == str(grafik_operasyon)
    ]
    if selected_rows.empty:
        st.info("Seçilen operasyon için veri bulunamadı.")
        return

    selected_op = selected_rows.iloc[0]
    yapilan = int(selected_op["yapilan_adet"])
    kalan = int(selected_op["kalan_adet"])

    gcol1, gcol2 = st.columns([1, 1])
    with gcol1:
        pie_data = pd.DataFrame({
            "Durum": ["Yapılan", "Kalan"],
            "Adet": [yapilan, kalan],
        })

        if int(pie_data["Adet"].sum()) <= 0:
            st.info("Bu operasyon için gösterilecek veri yok.")
        else:
            try:
                st.vega_lite_chart(
                    pie_data,
                    {
                        "mark": {"type": "arc", "innerRadius": 65},
                        "encoding": {
                            "theta": {"field": "Adet", "type": "quantitative"},
                            "color": {
                                "field": "Durum",
                                "type": "nominal",
                                "legend": {"title": "Durum"},
                            },
                            "tooltip": [
                                {"field": "Durum", "type": "nominal"},
                                {"field": "Adet", "type": "quantitative"},
                            ],
                        },
                        "view": {"stroke": None},
                    },
                    use_container_width=True,
                )
            except Exception as exc:
                st.warning("Daire grafiği gösterilemedi; veriler çubuk grafik olarak gösteriliyor.")
                st.bar_chart(pie_data.set_index("Durum"))

    with gcol2:
        st.metric("Yapılan", f"{yapilan} adet")
        st.metric("Kalan", f"{kalan} adet")
        st.metric("Tamamlanma", f"%{float(selected_op['tamamlanma_yuzdesi']):.1f}")
        st.info(
            "Bu grafik, seçilen operasyonda toplam takip edilen adet içinden "
            "kaç tanesinin yapıldığını ve kaç tanesinin kaldığını gösterir."
        )

def kayitlar_page():
    st.subheader("Kayıtlar")
    conn = sqlite3.connect(DB_PATH)
    sessions = pd.read_sql_query("SELECT * FROM work_sessions ORDER BY id DESC", conn)

    if sessions.empty:
        st.info("Henüz kayıt yok.")
        conn.close()
        return

    sessions["durum"] = sessions.apply(
        lambda row: production_status(
            int(row["siparis_adedi"]),
            int(row["saglam_tamamlanan"]),
        ),
        axis=1,
    )
    sessions["kalan"] = (
        sessions["siparis_adedi"] - sessions["saglam_tamamlanan"]
    ).clip(lower=0)
    st.dataframe(_rename_oc_column(sessions), use_container_width=True)

    selected_id = st.selectbox("Detay görmek için kayıt seç", sessions["id"].tolist())
    details = pd.read_sql_query(
        "SELECT urun_no, operasyon_sirasi, operasyon_adi, yapildi, fire_var, fire_operasyonu, fire_notu FROM operation_entries WHERE session_id = ? ORDER BY urun_no, operasyon_sirasi",
        conn,
        params=(int(selected_id),),
    )
    conn.close()

    st.write("Operasyon detayları")
    st.dataframe(details, use_container_width=True)

    selected_session = sessions[sessions["id"] == selected_id].iloc[0]
    render_operation_progress(details, selected_session, int(selected_id))

    st.divider()
    st.subheader("Güncelleme Geçmişi")
    history = get_update_history(int(selected_id))
    if history.empty:
        st.info("Bu kayıt henüz sonradan güncellenmemiş.")
    else:
        st.dataframe(history, use_container_width=True)

    st.info(
        "Bu kaydı tamamlamaya devam etmek için sol menüden "
        "'Üretime Devam Et' bölümünü açabilirsin."
    )

    st.divider()
    st.subheader("Kayıt Silme")
    st.warning(
        f"{selected_id} numaralı kayıt ve ona bağlı tüm operasyon detayları kalıcı olarak silinecek."
    )

    silme_onayi = st.checkbox(
        "Bu kaydı silmek istediğimi onaylıyorum",
        key=f"delete_confirm_{selected_id}",
    )

    if st.button(
        "Seçili Kaydı Sil",
        type="primary",
        disabled=not silme_onayi,
        key=f"delete_button_{selected_id}",
    ):
        delete_session(int(selected_id))
        st.success(f"{selected_id} numaralı kayıt silindi.")
        st.rerun()

    st.divider()
    csv = _rename_oc_column(sessions).to_csv(index=False).encode("utf-8-sig")
    st.download_button("Kayıtları CSV indir", csv, "uretim_kayitlari.csv", "text/csv")



def grafikler_page():
    st.subheader("Kombinasyon İlerleme Grafikleri")
    st.caption("Bir kayıt seçerek her operasyon için yapılan ve kalan adetleri görüntüleyebilirsin.")

    conn = sqlite3.connect(DB_PATH)
    sessions = pd.read_sql_query("SELECT * FROM work_sessions ORDER BY id DESC", conn)

    if sessions.empty:
        conn.close()
        st.info("Grafik göstermek için önce en az bir kayıt eklemelisin.")
        return

    labels = sessions.apply(
        lambda r: f"#{int(r['id'])} - {r['tarih']} - {r['operator_ismi']} - {r['pos']} - {r['kombinasyon_adi']}",
        axis=1,
    ).tolist()
    selected_label = st.selectbox("Grafiğini görmek istediğin kayıt", labels)
    selected_index = labels.index(selected_label)
    selected_session = sessions.iloc[selected_index]
    selected_id = int(selected_session["id"])

    details = pd.read_sql_query(
        "SELECT urun_no, operasyon_sirasi, operasyon_adi, yapildi, fire_var, fire_operasyonu, fire_notu "
        "FROM operation_entries WHERE session_id = ? ORDER BY urun_no, operasyon_sirasi",
        conn,
        params=(selected_id,),
    )
    conn.close()

    if details.empty:
        st.warning("Bu kaydın operasyon detayları yok. Grafik oluşturmak için yeni bir kayıt girerken operasyon kutularını işaretle.")
        return

    toplam_urun = int(selected_session.get("uretim_yuku", 0) or 0)
    summary = (
        details.groupby("operasyon_adi", as_index=False)["yapildi"]
        .sum()
        .rename(columns={"yapildi": "Yapılan"})
    )
    summary["Yapılan"] = pd.to_numeric(summary["Yapılan"], errors="coerce").fillna(0).astype(int)
    summary["Kalan"] = (toplam_urun - summary["Yapılan"]).clip(lower=0)
    summary["Tamamlanma %"] = ((summary["Yapılan"] / max(toplam_urun, 1)) * 100).round(1)

    m1, m2, m3 = st.columns(3)
    m1.metric("POS", selected_session["pos"])
    m2.metric("Toplam takip edilen adet", toplam_urun)
    m3.metric("Fire", int(selected_session["fire_adedi"]))

    st.write("Operasyon bazlı durum")
    st.dataframe(summary, use_container_width=True)

    st.write("Tüm operasyonlarda yapılan / kalan")
    st.bar_chart(summary.set_index("operasyon_adi")[["Yapılan", "Kalan"]])

    operasyon = st.selectbox("Daire grafiği için operasyon seç", summary["operasyon_adi"].tolist())
    row = summary[summary["operasyon_adi"] == operasyon].iloc[0]
    pie_data = pd.DataFrame({"Durum": ["Yapılan", "Kalan"], "Adet": [int(row["Yapılan"]), int(row["Kalan"])]})

    if pie_data["Adet"].sum() == 0:
        st.info("Bu operasyon için gösterilecek adet yok.")
        return

    st.vega_lite_chart(
        pie_data,
        {
            "mark": {"type": "arc", "innerRadius": 70},
            "encoding": {
                "theta": {"field": "Adet", "type": "quantitative"},
                "color": {"field": "Durum", "type": "nominal", "legend": {"title": "Durum"}},
                "tooltip": [
                    {"field": "Durum", "type": "nominal"},
                    {"field": "Adet", "type": "quantitative"},
                ],
            },
            "view": {"stroke": None},
        },
        use_container_width=True,
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Yapılan", int(row["Yapılan"]))
    c2.metric("Kalan", int(row["Kalan"]))
    c3.metric("Tamamlanma", f"%{float(row['Tamamlanma %']):.1f}")


def build_manager_excel_report(
    filtered_sessions: pd.DataFrame,
    by_operator: pd.DataFrame,
    by_pos: pd.DataFrame,
    daily: pd.DataFrame,
    fire_by_operation: pd.DataFrame,
) -> bytes:
    """Yönetici panelindeki tabloları tek Excel raporunda toplar."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        _rename_oc_column(filtered_sessions).to_excel(writer, sheet_name="Filtreli Kayıtlar", index=False)
        _rename_oc_column(by_operator).to_excel(writer, sheet_name="İşçi Özeti", index=False)
        _rename_oc_column(by_pos).to_excel(writer, sheet_name="POS Özeti", index=False)
        _rename_oc_column(daily).to_excel(writer, sheet_name="Günlük Özet", index=False)
        _rename_oc_column(fire_by_operation).to_excel(writer, sheet_name="Fire Operasyonları", index=False)
    return output.getvalue()


def yonetici_paneli_page():
    st.subheader("Yönetici Paneli")
    st.caption(
        "Tarih, işçi, OC ve POS filtreleriyle üretim, fire, çalışma süresi ve verimlilik durumunu gösterir."
    )

    conn = sqlite3.connect(DB_PATH)
    sessions = pd.read_sql_query(
        "SELECT * FROM work_sessions ORDER BY tarih, id",
        conn,
    )
    operations = pd.read_sql_query(
        """
        SELECT
            session_id,
            urun_no,
            operasyon_adi,
            yapildi,
            fire_var,
            fire_operasyonu,
            fire_notu
        FROM operation_entries
        """,
        conn,
    )
    conn.close()

    if sessions.empty:
        st.info("Yönetici paneli için henüz kayıt yok.")
        return

    sessions["tarih"] = pd.to_datetime(sessions["tarih"], errors="coerce")
    sessions = sessions.dropna(subset=["tarih"]).copy()

    if sessions.empty:
        st.warning("Kayıtlardaki tarih bilgileri okunamadı.")
        return

    min_date = sessions["tarih"].min().date()
    max_date = sessions["tarih"].max().date()

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        tarih_araligi = st.date_input(
            "Tarih aralığı",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            key="manager_date_range",
        )
    with f2:
        operator_options = ["Tümü"] + sorted(
            sessions["operator_ismi"].dropna().astype(str).unique().tolist()
        )
        selected_operator = st.selectbox(
            "İşçi / Operatör",
            operator_options,
            key="manager_operator",
        )
    with f3:
        project_options = ["Tümü"] + sorted(
            value
            for value in sessions["proje"].fillna("").astype(str).unique().tolist()
            if value.strip()
        )
        selected_project = st.selectbox(
            "OC",
            project_options,
            key="manager_project",
        )
    with f4:
        pos_values = sessions["pos"].dropna().astype(str).unique().tolist()
        pos_values = sorted(
            pos_values,
            key=lambda value: int("".join(filter(str.isdigit, value)) or 0),
        )
        selected_pos = st.selectbox(
            "POS",
            ["Tümü"] + pos_values,
            key="manager_pos",
        )

    if isinstance(tarih_araligi, (tuple, list)) and len(tarih_araligi) == 2:
        start_date, end_date = tarih_araligi
    else:
        start_date = end_date = tarih_araligi

    mask = (
        (sessions["tarih"].dt.date >= start_date)
        & (sessions["tarih"].dt.date <= end_date)
    )

    if selected_operator != "Tümü":
        mask &= sessions["operator_ismi"].astype(str) == selected_operator
    if selected_project != "Tümü":
        mask &= sessions["proje"].fillna("").astype(str) == selected_project
    if selected_pos != "Tümü":
        mask &= sessions["pos"].astype(str) == selected_pos

    filtered = sessions.loc[mask].copy()

    if filtered.empty:
        st.warning("Seçilen filtrelere uygun kayıt bulunamadı.")
        return

    overtime_filtered = filtered[
        filtered["calisma_tipi"] == "Mesaili"
    ].copy()
    if not overtime_filtered.empty:
        overtime_filtered["tarih_yazi"] = (
            overtime_filtered["tarih"].dt.date.astype(str)
        )
        overtime_summary = (
            overtime_filtered.groupby(
                ["tarih_yazi", "operator_ismi"],
                as_index=False,
            )
            .agg(
                toplam_saat=("calisma_saati", "sum"),
                uretilen_adet=("saglam_tamamlanan", "sum"),
                fire_adedi=("fire_adedi", "sum"),
                POS=("pos", lambda values: ", ".join(sorted(set(map(str, values))))),
                OC=("proje", lambda values: ", ".join(sorted(set(map(str, values))))),
            )
        )
        st.warning(
            f"Seçilen tarih aralığında {len(overtime_summary)} mesai kaydı var."
        )
        with st.expander("Mesai kayıtlarını göster", expanded=False):
            st.dataframe(
                overtime_summary.rename(
                    columns={
                        "tarih_yazi": "Tarih",
                        "operator_ismi": "İşçi",
                        "toplam_saat": "Toplam Saat",
                        "uretilen_adet": "Üretilen",
                        "fire_adedi": "Fire",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    total_records = int(len(filtered))
    total_order = int(filtered["siparis_adedi"].sum())
    total_completed = int(filtered["saglam_tamamlanan"].sum())
    total_fire = int(filtered["fire_adedi"].sum())
    total_hours = float(filtered["calisma_saati"].sum())
    total_load = int(filtered["uretim_yuku"].sum())
    remaining = max(total_order - total_completed, 0)
    completion_rate = (total_completed / total_order * 100) if total_order else 0.0
    fire_rate = (total_fire / total_load * 100) if total_load else 0.0
    productivity = (total_completed / total_hours) if total_hours else 0.0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Kayıt sayısı", total_records)
    k2.metric("Toplam sipariş", total_order)
    k3.metric("Sağlam tamamlanan", total_completed)
    k4.metric("Kalan sipariş", remaining)

    k5, k6, k7, k8 = st.columns(4)
    k5.metric("Toplam fire", total_fire)
    k6.metric("Fire oranı", f"%{fire_rate:.1f}")
    k7.metric("Tamamlanma oranı", f"%{completion_rate:.1f}")
    k8.metric("Verimlilik", f"{productivity:.2f} adet/saat")

    st.divider()

    by_operator = (
        filtered.groupby("operator_ismi", as_index=False)
        .agg(
            kayıt_sayısı=("id", "count"),
            çalışma_saati=("calisma_saati", "sum"),
            sipariş=("siparis_adedi", "sum"),
            tamamlanan=("saglam_tamamlanan", "sum"),
            fire=("fire_adedi", "sum"),
        )
    )
    by_operator["kalan"] = (
        by_operator["sipariş"] - by_operator["tamamlanan"]
    ).clip(lower=0)
    by_operator["fire_oranı_%"] = (
        by_operator["fire"]
        / (by_operator["tamamlanan"] + by_operator["fire"]).replace(0, 1)
        * 100
    ).round(1)
    by_operator["adet_saat"] = (
        by_operator["tamamlanan"]
        / by_operator["çalışma_saati"].replace(0, 1)
    ).round(2)
    by_operator = by_operator.sort_values(
        ["tamamlanan", "adet_saat"],
        ascending=[False, False],
    )

    by_pos = (
        filtered.groupby("pos", as_index=False)
        .agg(
            kayıt_sayısı=("id", "count"),
            sipariş=("siparis_adedi", "sum"),
            tamamlanan=("saglam_tamamlanan", "sum"),
            fire=("fire_adedi", "sum"),
            üretim_yükü=("uretim_yuku", "sum"),
            çalışma_saati=("calisma_saati", "sum"),
        )
    )
    by_pos["kalan"] = (by_pos["sipariş"] - by_pos["tamamlanan"]).clip(lower=0)
    by_pos["fire_oranı_%"] = (
        by_pos["fire"] / by_pos["üretim_yükü"].replace(0, 1) * 100
    ).round(1)
    by_pos["adet_saat"] = (
        by_pos["tamamlanan"] / by_pos["çalışma_saati"].replace(0, 1)
    ).round(2)
    by_pos["_pos_no"] = (
        by_pos["pos"].astype(str).str.extract(r"(\d+)")[0].fillna(0).astype(int)
    )
    by_pos = by_pos.sort_values("_pos_no").drop(columns=["_pos_no"])

    filtered["tarih_gün"] = filtered["tarih"].dt.date.astype(str)
    daily = (
        filtered.groupby("tarih_gün", as_index=False)
        .agg(
            sipariş=("siparis_adedi", "sum"),
            tamamlanan=("saglam_tamamlanan", "sum"),
            fire=("fire_adedi", "sum"),
            çalışma_saati=("calisma_saati", "sum"),
        )
        .sort_values("tarih_gün")
    )

    filtered_ids = set(filtered["id"].astype(int).tolist())
    selected_operations = operations[
        operations["session_id"].astype(int).isin(filtered_ids)
    ].copy()

    fire_rows = selected_operations[
        (selected_operations["fire_var"] == 1)
        & selected_operations["fire_operasyonu"]
        .fillna("")
        .astype(str)
        .str.strip()
        .ne("")
    ].copy()

    if fire_rows.empty:
        fire_by_operation = pd.DataFrame(
            columns=["fire_operasyonu", "fire_adedi"]
        )
    else:
        fire_by_operation = (
            fire_rows.drop_duplicates(["session_id", "urun_no"])
            .groupby("fire_operasyonu", as_index=False)
            .size()
            .rename(columns={"size": "fire_adedi"})
            .sort_values("fire_adedi", ascending=False)
        )

    tab1, tab2, tab3, tab4 = st.tabs(
        ["İşçi Performansı", "POS Performansı", "Günlük Gidişat", "Fire Analizi"]
    )

    with tab1:
        st.write("İşçi / operatör bazlı performans")
        st.dataframe(by_operator, use_container_width=True)
        if not by_operator.empty:
            st.bar_chart(
                by_operator.set_index("operator_ismi")[["tamamlanan", "fire"]]
            )

    with tab2:
        st.write("POS bazlı üretim durumu")
        st.dataframe(by_pos, use_container_width=True)
        if not by_pos.empty:
            st.bar_chart(
                by_pos.set_index("pos")[["sipariş", "tamamlanan", "fire"]]
            )

    with tab3:
        st.write("Günlük üretim ve fire")
        st.dataframe(daily, use_container_width=True)
        if not daily.empty:
            st.line_chart(
                daily.set_index("tarih_gün")[["sipariş", "tamamlanan", "fire"]]
            )

    with tab4:
        st.write("Fire oluşan operasyonlar")
        if fire_by_operation.empty:
            st.info("Seçilen filtrelerde operasyon bazlı fire kaydı yok.")
        else:
            st.dataframe(fire_by_operation, use_container_width=True)
            st.bar_chart(
                fire_by_operation.set_index("fire_operasyonu")[["fire_adedi"]]
            )

    st.divider()
    report_source = filtered.drop(columns=["tarih_gün"], errors="ignore").copy()
    report_source["tarih"] = report_source["tarih"].dt.strftime("%Y-%m-%d")

    report_bytes = build_manager_excel_report(
        filtered_sessions=report_source,
        by_operator=by_operator,
        by_pos=by_pos,
        daily=daily,
        fire_by_operation=fire_by_operation,
    )
    st.download_button(
        "Yönetici Excel Raporunu İndir",
        data=report_bytes,
        file_name=f"yonetici_uretim_raporu_{start_date}_{end_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


# -----------------------------
# GÜN SONU KONTROLÜ VE RAPORLAMA
# -----------------------------
def _pdf_font_name() -> str:
    """Türkçe karakterleri destekleyen bir font bulur; bulunamazsa Helvetica kullanır."""
    if "DejaVuSans" in pdfmetrics.getRegisteredFontNames():
        return "DejaVuSans"

    candidates = [
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/dejavu/DejaVuSans.ttf"),
        Path("C:/Windows/Fonts/arial.ttf"),
    ]
    for font_path in candidates:
        if font_path.exists():
            try:
                pdfmetrics.registerFont(TTFont("DejaVuSans", str(font_path)))
                return "DejaVuSans"
            except Exception:
                continue
    return "Helvetica"


def _pdf_safe_text(value, font_name: str) -> str:
    text_value = "" if pd.isna(value) else str(value)
    if font_name == "Helvetica":
        return (
            unicodedata.normalize("NFKD", text_value)
            .encode("ascii", "ignore")
            .decode("ascii")
        )
    return text_value


def _table_for_pdf(df: pd.DataFrame, font_name: str, max_rows: int = 60):
    """DataFrame'i PDF'e uygun, satır tekrar başlıklı tabloya çevirir."""
    if df.empty:
        return Paragraph(_pdf_safe_text("Kayıt bulunamadı.", font_name), ParagraphStyle(
            "Empty", fontName=font_name, fontSize=9, leading=12
        ))

    work = _rename_oc_column(df.head(max_rows).copy())
    for column in work.columns:
        work[column] = work[column].map(lambda x: _pdf_safe_text(x, font_name))

    headers = [
        Paragraph(f"<b>{_pdf_safe_text(col, font_name)}</b>", ParagraphStyle(
            f"Header_{col}", fontName=font_name, fontSize=7, leading=8
        ))
        for col in work.columns
    ]
    rows = []
    body_style = ParagraphStyle("BodyCell", fontName=font_name, fontSize=6.5, leading=8)
    for _, row in work.iterrows():
        rows.append([Paragraph(_pdf_safe_text(value, font_name), body_style) for value in row])

    available_width = 272 * mm
    col_width = available_width / max(len(work.columns), 1)
    table = Table([headers] + rows, repeatRows=1, colWidths=[col_width] * len(work.columns))
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#12344D")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#AAB7C4")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FB")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return table


def build_day_end_pdf(
    selected_date,
    metrics: dict,
    alerts: list,
    worker_summary: pd.DataFrame,
    pos_summary: pd.DataFrame,
    incomplete_ops: pd.DataFrame,
    fire_details: pd.DataFrame,
    part_time: pd.DataFrame,
    untouched_units: pd.DataFrame,
) -> bytes:
    """Gün sonu kontrol verilerini yatay A4 PDF raporuna dönüştürür."""
    output = BytesIO()
    font_name = _pdf_font_name()
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "DayEndTitle",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#12344D"),
    )
    heading_style = ParagraphStyle(
        "DayEndHeading",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=12,
        leading=15,
        textColor=colors.HexColor("#12344D"),
        spaceBefore=8,
        spaceAfter=5,
    )
    body_style = ParagraphStyle(
        "DayEndBody",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=9,
        leading=12,
    )

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=f"Gün Sonu Üretim Raporu - {selected_date}",
    )

    story = [
        Paragraph(_pdf_safe_text("Gün Sonu Üretim Kontrol Raporu", font_name), title_style),
        Paragraph(_pdf_safe_text(f"Tarih: {selected_date}", font_name), body_style),
        Spacer(1, 5 * mm),
    ]

    metric_df = pd.DataFrame([
        {
            "Kayıt": metrics["record_count"],
            "İşçi": metrics["worker_count"],
            "Sipariş": metrics["order"],
            "Tamamlanan": metrics["completed"],
            "Kalan": metrics["remaining"],
            "Fire": metrics["fire"],
            "Fire %": f"{metrics['fire_rate']:.1f}",
            "Çalışma saati": f"{metrics['hours']:.1f}",
            "Adet/saat": f"{metrics['productivity']:.2f}",
        }
    ])
    story.append(_table_for_pdf(metric_df, font_name, max_rows=5))

    story.append(Paragraph(_pdf_safe_text("Otomatik Uyarılar", font_name), heading_style))
    if alerts:
        for alert in alerts:
            story.append(Paragraph(
                _pdf_safe_text(f"- {alert['level'].upper()}: {alert['message']}", font_name),
                body_style,
            ))
    else:
        story.append(Paragraph(_pdf_safe_text("Kritik uyarı bulunmadı.", font_name), body_style))

    sections = [
        ("İşçi Özeti", worker_summary),
        ("POS Özeti", pos_summary),
        ("Eksik Operasyonlar", incomplete_ops),
        ("Fire Ayrıntıları", fire_details),
        ("Yarı Zamanlı Çalışanlar", part_time),
        ("Hiç Operasyon İşaretlenmemiş Ürünler", untouched_units),
    ]

    for index, (title, df) in enumerate(sections):
        story.append(PageBreak())
        story.append(Paragraph(_pdf_safe_text(title, font_name), heading_style))
        story.append(_table_for_pdf(df, font_name))

    doc.build(story)
    return output.getvalue()


def build_day_end_excel(
    sessions_day: pd.DataFrame,
    worker_summary: pd.DataFrame,
    pos_summary: pd.DataFrame,
    incomplete_ops: pd.DataFrame,
    fire_details: pd.DataFrame,
    part_time: pd.DataFrame,
    untouched_units: pd.DataFrame,
    alerts: list,
) -> bytes:
    output = BytesIO()
    alerts_df = pd.DataFrame(alerts)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        _rename_oc_column(sessions_day).to_excel(writer, sheet_name="Günlük Kayıtlar", index=False)
        _rename_oc_column(worker_summary).to_excel(writer, sheet_name="İşçi Özeti", index=False)
        _rename_oc_column(pos_summary).to_excel(writer, sheet_name="POS Özeti", index=False)
        _rename_oc_column(incomplete_ops).to_excel(writer, sheet_name="Eksik Operasyonlar", index=False)
        _rename_oc_column(fire_details).to_excel(writer, sheet_name="Fire Ayrıntıları", index=False)
        _rename_oc_column(part_time).to_excel(writer, sheet_name="Yarı Zamanlı", index=False)
        _rename_oc_column(untouched_units).to_excel(writer, sheet_name="İşlemsiz Ürünler", index=False)
        alerts_df.to_excel(writer, sheet_name="Uyarılar", index=False)
    return output.getvalue()


def gun_sonu_kontrolu_page():
    st.subheader("Gün Sonu Kontrolü")
    st.caption(
        "Seçilen günün tamamlanmayan POS'larını, eksik operasyonlarını, firelerini, "
        "yarı zamanlı çalışanlarını ve otomatik uyarılarını tek ekranda gösterir."
    )

    conn = sqlite3.connect(DB_PATH)
    sessions = pd.read_sql_query(
        "SELECT * FROM work_sessions ORDER BY tarih, id",
        conn,
    )
    operations = pd.read_sql_query(
        """
        SELECT
            session_id, urun_no, operasyon_sirasi, operasyon_adi,
            yapildi, fire_var, fire_operasyonu, fire_notu
        FROM operation_entries
        ORDER BY session_id, urun_no, operasyon_sirasi
        """,
        conn,
    )
    conn.close()

    if sessions.empty:
        st.info("Gün sonu kontrolü için henüz kayıt yok.")
        return

    sessions["tarih_dt"] = pd.to_datetime(sessions["tarih"], errors="coerce")
    sessions = sessions.dropna(subset=["tarih_dt"]).copy()
    if sessions.empty:
        st.warning("Kayıtlardaki tarih bilgileri okunamadı.")
        return

    available_dates = sorted(sessions["tarih_dt"].dt.date.unique(), reverse=True)
    default_date = available_dates[0]

    d1, d2, d3 = st.columns(3)
    with d1:
        selected_date = st.date_input(
            "Kontrol tarihi",
            value=default_date,
            min_value=min(available_dates),
            max_value=max(available_dates),
            key="day_end_date",
        )
    with d2:
        fire_threshold = st.number_input(
            "Fire uyarı sınırı (%)",
            min_value=0.0,
            max_value=100.0,
            value=10.0,
            step=1.0,
            key="day_end_fire_threshold",
        )
    with d3:
        low_productivity_threshold = st.number_input(
            "Düşük verimlilik sınırı (adet/saat)",
            min_value=0.0,
            value=1.0,
            step=0.25,
            key="day_end_productivity_threshold",
        )

    day_mask = sessions["tarih_dt"].dt.date == selected_date
    sessions_day = sessions.loc[day_mask].copy()
    if sessions_day.empty:
        st.warning("Seçilen tarihte kayıt bulunamadı.")
        return

    sessions_day["kalan"] = (
        sessions_day["siparis_adedi"] - sessions_day["saglam_tamamlanan"]
    ).clip(lower=0)
    sessions_day["tamamlanma_%"] = (
        sessions_day["saglam_tamamlanan"]
        / sessions_day["siparis_adedi"].replace(0, 1)
        * 100
    ).round(1)
    sessions_day["fire_%"] = (
        sessions_day["fire_adedi"]
        / sessions_day["uretim_yuku"].replace(0, 1)
        * 100
    ).round(1)
    sessions_day["adet_saat"] = (
        sessions_day["saglam_tamamlanan"]
        / sessions_day["calisma_saati"].replace(0, 1)
    ).round(2)

    selected_ids = set(sessions_day["id"].astype(int).tolist())
    operations_day = operations[
        operations["session_id"].astype(int).isin(selected_ids)
    ].copy()

    worker_summary = (
        sessions_day.groupby("operator_ismi", as_index=False)
        .agg(
            kayıt=("id", "count"),
            çalışma_saati=("calisma_saati", "sum"),
            sipariş=("siparis_adedi", "sum"),
            tamamlanan=("saglam_tamamlanan", "sum"),
            fire=("fire_adedi", "sum"),
        )
    )
    worker_summary["kalan"] = (
        worker_summary["sipariş"] - worker_summary["tamamlanan"]
    ).clip(lower=0)
    worker_summary["adet_saat"] = (
        worker_summary["tamamlanan"]
        / worker_summary["çalışma_saati"].replace(0, 1)
    ).round(2)
    worker_summary["fire_%"] = (
        worker_summary["fire"]
        / (worker_summary["tamamlanan"] + worker_summary["fire"]).replace(0, 1)
        * 100
    ).round(1)
    worker_summary = worker_summary.sort_values(
        ["tamamlanan", "adet_saat"], ascending=[False, False]
    )

    pos_summary = (
        sessions_day.groupby(["pos", "proje", "kombinasyon_adi"], dropna=False, as_index=False)
        .agg(
            kayıt=("id", "count"),
            sipariş=("siparis_adedi", "sum"),
            tamamlanan=("saglam_tamamlanan", "sum"),
            fire=("fire_adedi", "sum"),
            üretim_yükü=("uretim_yuku", "sum"),
            çalışma_saati=("calisma_saati", "sum"),
        )
    )
    pos_summary["kalan"] = (
        pos_summary["sipariş"] - pos_summary["tamamlanan"]
    ).clip(lower=0)
    pos_summary["tamamlanma_%"] = (
        pos_summary["tamamlanan"] / pos_summary["sipariş"].replace(0, 1) * 100
    ).round(1)
    pos_summary["fire_%"] = (
        pos_summary["fire"] / pos_summary["üretim_yükü"].replace(0, 1) * 100
    ).round(1)
    pos_summary["adet_saat"] = (
        pos_summary["tamamlanan"] / pos_summary["çalışma_saati"].replace(0, 1)
    ).round(2)
    pos_summary["_pos_no"] = (
        pos_summary["pos"].astype(str).str.extract(r"(\d+)")[0].fillna(0).astype(int)
    )
    pos_summary = pos_summary.sort_values("_pos_no").drop(columns=["_pos_no"])

    session_lookup = sessions_day[
        ["id", "operator_ismi", "proje", "pos", "kombinasyon_adi", "uretim_yuku"]
    ].rename(columns={"id": "session_id"})

    if operations_day.empty:
        incomplete_ops = pd.DataFrame(columns=[
            "session_id", "operator_ismi", "proje", "pos", "kombinasyon_adi",
            "operasyon_adi", "yapılan", "kalan"
        ])
        fire_details = pd.DataFrame(columns=[
            "session_id", "operator_ismi", "proje", "pos", "urun_no",
            "fire_operasyonu", "fire_notu"
        ])
        untouched_units = pd.DataFrame(columns=[
            "session_id", "operator_ismi", "proje", "pos", "urun_no"
        ])
    else:
        operation_progress = (
            operations_day.groupby(["session_id", "operasyon_adi"], as_index=False)
            .agg(yapılan=("yapildi", "sum"))
            .merge(session_lookup, on="session_id", how="left")
        )
        operation_progress["kalan"] = (
            operation_progress["uretim_yuku"] - operation_progress["yapılan"]
        ).clip(lower=0)
        incomplete_ops = operation_progress[
            operation_progress["kalan"] > 0
        ][[
            "session_id", "operator_ismi", "proje", "pos", "kombinasyon_adi",
            "operasyon_adi", "yapılan", "kalan"
        ]].sort_values(["pos", "session_id", "operasyon_adi"])

        fire_rows = operations_day[
            operations_day["fire_var"] == 1
        ].drop_duplicates(["session_id", "urun_no"]).copy()
        fire_details = fire_rows.merge(session_lookup, on="session_id", how="left")
        if fire_details.empty:
            fire_details = pd.DataFrame(columns=[
                "session_id", "operator_ismi", "proje", "pos", "urun_no",
                "fire_operasyonu", "fire_notu"
            ])
        else:
            fire_details = fire_details[[
                "session_id", "operator_ismi", "proje", "pos", "urun_no",
                "fire_operasyonu", "fire_notu"
            ]].sort_values(["pos", "session_id", "urun_no"])

        unit_progress = (
            operations_day.groupby(["session_id", "urun_no"], as_index=False)
            .agg(yapılan_operasyon=("yapildi", "sum"))
        )
        untouched_units = unit_progress[
            unit_progress["yapılan_operasyon"] == 0
        ].merge(session_lookup, on="session_id", how="left")
        untouched_units = untouched_units[[
            "session_id", "operator_ismi", "proje", "pos", "kombinasyon_adi",
            "urun_no", "yapılan_operasyon"
        ]].sort_values(["pos", "session_id", "urun_no"])

    part_time = sessions_day[
        sessions_day["calisma_tipi"] == "Yarı zamanlı"
    ][[
        "id", "operator_ismi", "proje", "pos", "calisma_saati", "neden"
    ]].rename(columns={"id": "kayıt_no"})

    overtime_day = (
        sessions_day[
            sessions_day["calisma_tipi"] == "Mesaili"
        ]
        .groupby("operator_ismi", as_index=False)
        .agg(
            toplam_saat=("calisma_saati", "sum"),
            uretilen_adet=("saglam_tamamlanan", "sum"),
            fire_adedi=("fire_adedi", "sum"),
            POS=("pos", lambda values: ", ".join(sorted(set(map(str, values))))),
            OC=("proje", lambda values: ", ".join(sorted(set(map(str, values))))),
        )
    )

    total_order = int(sessions_day["siparis_adedi"].sum())
    total_completed = int(sessions_day["saglam_tamamlanan"].sum())
    total_fire = int(sessions_day["fire_adedi"].sum())
    total_load = int(sessions_day["uretim_yuku"].sum())
    total_hours = float(sessions_day["calisma_saati"].sum())
    remaining = max(total_order - total_completed, 0)
    fire_rate = (total_fire / total_load * 100) if total_load else 0.0
    productivity = (total_completed / total_hours) if total_hours else 0.0

    metrics = {
        "record_count": int(len(sessions_day)),
        "worker_count": int(sessions_day["operator_ismi"].nunique()),
        "order": total_order,
        "completed": total_completed,
        "remaining": remaining,
        "fire": total_fire,
        "fire_rate": fire_rate,
        "hours": total_hours,
        "productivity": productivity,
    }

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Çalışan işçi", metrics["worker_count"])
    m2.metric("Toplam sipariş", total_order)
    m3.metric("Tamamlanan", total_completed)
    m4.metric("Kalan", remaining)

    m5, m6, m7, m8 = st.columns(4)
    m5.metric("Fire", total_fire)
    m6.metric("Fire oranı", f"%{fire_rate:.1f}")
    m7.metric("Çalışma saati", f"{total_hours:.1f}")
    m8.metric("Verimlilik", f"{productivity:.2f} adet/saat")

    alerts = []
    for _, row in pos_summary.iterrows():
        if float(row["fire_%"]) >= float(fire_threshold) and int(row["fire"]) > 0:
            alerts.append({
                "level": "kritik",
                "message": f"{row['pos']} fire oranı %{float(row['fire_%']):.1f}; sınır %{fire_threshold:.1f}.",
            })
        if int(row["kalan"]) > 0:
            alerts.append({
                "level": "uyarı",
                "message": f"{row['pos']} için {int(row['kalan'])} adet sipariş tamamlanmadı.",
            })
        else:
            alerts.append({
                "level": "tamam",
                "message": f"{row['pos']} siparişi tamamlandı.",
            })

    for _, row in worker_summary.iterrows():
        if float(row["adet_saat"]) < float(low_productivity_threshold):
            alerts.append({
                "level": "uyarı",
                "message": f"{row['operator_ismi']} verimliliği {float(row['adet_saat']):.2f} adet/saat; sınır {low_productivity_threshold:.2f}.",
            })

    for _, row in part_time.iterrows():
        alerts.append({
            "level": "bilgi",
            "message": f"{row['operator_ismi']} yarı zamanlı çalıştı ({float(row['calisma_saati']):.1f} saat). Neden: {row['neden'] or '-'}.",
        })

    for _, row in overtime_day.iterrows():
        alerts.append({
            "level": "uyarı",
            "message": (
                f"MESAİ: {row['operator_ismi']} toplam "
                f"{float(row['toplam_saat']):.1f} saat çalıştı. "
                f"POS: {row['POS']}."
            ),
        })

    if not untouched_units.empty:
        alerts.append({
            "level": "kritik",
            "message": f"{len(untouched_units)} ürün için hiçbir operasyon işaretlenmedi.",
        })
    if not incomplete_ops.empty:
        alerts.append({
            "level": "uyarı",
            "message": f"{len(incomplete_ops)} operasyon satırında tamamlanmamış iş bulunuyor.",
        })

    st.divider()
    st.subheader("Otomatik Uyarılar")
    if not alerts:
        st.success("Seçilen gün için uyarı bulunmadı.")
    else:
        for alert in alerts:
            if alert["level"] == "kritik":
                st.error(alert["message"])
            elif alert["level"] == "uyarı":
                st.warning(alert["message"])
            elif alert["level"] == "tamam":
                st.success(alert["message"])
            else:
                st.info(alert["message"])

    tabs = st.tabs([
        "POS Durumu", "Eksik Operasyonlar", "Fireler",
        "İşçi Durumu", "İşlemsiz Ürünler"
    ])

    with tabs[0]:
        st.dataframe(_rename_oc_column(pos_summary), use_container_width=True)
        st.bar_chart(pos_summary.set_index("pos")[["sipariş", "tamamlanan", "fire"]])

    with tabs[1]:
        if incomplete_ops.empty:
            st.success("Eksik operasyon bulunmuyor.")
        else:
            st.dataframe(_rename_oc_column(incomplete_ops), use_container_width=True)

    with tabs[2]:
        if fire_details.empty:
            st.success("Fire kaydı bulunmuyor.")
        else:
            st.dataframe(_rename_oc_column(fire_details), use_container_width=True)

    with tabs[3]:
        st.write("İşçi özeti")
        st.dataframe(worker_summary, use_container_width=True)
        st.write("Yarı zamanlı çalışanlar")
        if part_time.empty:
            st.success("Yarı zamanlı çalışma kaydı bulunmuyor.")
        else:
            st.dataframe(_rename_oc_column(part_time), use_container_width=True)

        st.write("Mesaili çalışanlar")
        if overtime_day.empty:
            st.success("Mesaili çalışma kaydı bulunmuyor.")
        else:
            st.dataframe(
                overtime_day.rename(
                    columns={
                        "operator_ismi": "İşçi",
                        "toplam_saat": "Toplam Saat",
                        "uretilen_adet": "Üretilen",
                        "fire_adedi": "Fire",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    with tabs[4]:
        if untouched_units.empty:
            st.success("Operasyon işareti olmayan ürün bulunmuyor.")
        else:
            st.dataframe(_rename_oc_column(untouched_units), use_container_width=True)

    st.divider()
    report_sessions = sessions_day.drop(columns=["tarih_dt"], errors="ignore").copy()
    excel_bytes = build_day_end_excel(
        sessions_day=report_sessions,
        worker_summary=worker_summary,
        pos_summary=pos_summary,
        incomplete_ops=incomplete_ops,
        fire_details=fire_details,
        part_time=part_time,
        untouched_units=untouched_units,
        alerts=alerts,
    )
    pdf_bytes = build_day_end_pdf(
        selected_date=selected_date,
        metrics=metrics,
        alerts=alerts,
        worker_summary=worker_summary,
        pos_summary=pos_summary,
        incomplete_ops=incomplete_ops,
        fire_details=fire_details,
        part_time=part_time,
        untouched_units=untouched_units,
    )

    b1, b2 = st.columns(2)
    with b1:
        st.download_button(
            "Gün Sonu Excel Raporunu İndir",
            data=excel_bytes,
            file_name=f"gun_sonu_raporu_{selected_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    with b2:
        st.download_button(
            "Gün Sonu PDF Raporunu İndir",
            data=pdf_bytes,
            file_name=f"gun_sonu_raporu_{selected_date}.pdf",
            mime="application/pdf",
            type="primary",
        )



def ozet_page():
    st.subheader("Özet")
    conn = sqlite3.connect(DB_PATH)
    sessions = pd.read_sql_query("SELECT * FROM work_sessions", conn)
    conn.close()

    if sessions.empty:
        st.info("Özet için kayıt yok.")
        return

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Toplam sipariş", int(sessions["siparis_adedi"].sum()))
    col2.metric("Sağlam tamamlanan", int(sessions["saglam_tamamlanan"].sum()))
    col3.metric("Fire", int(sessions["fire_adedi"].sum()))
    col4.metric("Üretim yükü", int(sessions["uretim_yuku"].sum()))

    by_pos = sessions.groupby("pos", as_index=False).agg(
        siparis_adedi=("siparis_adedi", "sum"),
        saglam_tamamlanan=("saglam_tamamlanan", "sum"),
        fire_adedi=("fire_adedi", "sum"),
        uretim_yuku=("uretim_yuku", "sum"),
    )

    st.write("POS bazlı özet")
    st.dataframe(by_pos, use_container_width=True)
    st.bar_chart(by_pos.set_index("pos")[["siparis_adedi", "saglam_tamamlanan", "fire_adedi"]])


if __name__ == "__main__":
    main()
