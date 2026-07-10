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
            source_type TEXT NOT NULL DEFAULT 'legacy',
            source_ref INTEGER,
            FOREIGN KEY(item_id) REFERENCES production_output_items(id)
        )
        """
    )

    cur.execute("PRAGMA table_info(production_output_items)")
    output_item_columns = {row[1] for row in cur.fetchall()}
    if "combination_name" not in output_item_columns:
        cur.execute(
            "ALTER TABLE production_output_items "
            "ADD COLUMN combination_name TEXT NOT NULL DEFAULT ''"
        )

    cur.execute("PRAGMA table_info(production_output_logs)")
    output_log_columns = {row[1] for row in cur.fetchall()}
    if "source_type" not in output_log_columns:
        cur.execute(
            "ALTER TABLE production_output_logs "
            "ADD COLUMN source_type TEXT NOT NULL DEFAULT 'legacy'"
        )
    if "source_ref" not in output_log_columns:
        cur.execute(
            "ALTER TABLE production_output_logs "
            "ADD COLUMN source_ref INTEGER"
        )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS operation_batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tarih TEXT NOT NULL,
            operator_ismi TEXT NOT NULL,
            calisma_tipi TEXT NOT NULL,
            calisma_saati REAL NOT NULL,
            neden TEXT,
            notlar TEXT,
            created_at TEXT NOT NULL
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS operation_work_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL,
            item_id INTEGER NOT NULL,
            operation_name TEXT NOT NULL,
            processed_qty INTEGER NOT NULL,
            fire_qty INTEGER NOT NULL DEFAULT 0,
            good_qty INTEGER NOT NULL,
            fire_note TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(batch_id) REFERENCES operation_batches(id),
            FOREIGN KEY(item_id) REFERENCES production_output_items(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS pos_operation_plan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id INTEGER NOT NULL,
            combination_name TEXT NOT NULL,
            operation_order INTEGER NOT NULL,
            operation_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(item_id, operation_name),
            FOREIGN KEY(item_id) REFERENCES production_output_items(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS worker_competencies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            worker_name TEXT NOT NULL,
            operation_name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(worker_name, operation_name)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS operation_performance_targets (
            operation_name TEXT PRIMARY KEY,
            target_qty_per_hour REAL NOT NULL DEFAULT 0,
            target_area_per_hour REAL NOT NULL DEFAULT 0,
            fire_limit_pct REAL NOT NULL DEFAULT 5,
            slow_limit_pct REAL NOT NULL DEFAULT 80,
            fast_limit_pct REAL NOT NULL DEFAULT 120,
            updated_at TEXT NOT NULL
        )
        """
    )

    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_operation_logs_item "
        "ON operation_work_logs(item_id, operation_name)"
    )
    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_operation_logs_batch "
        "ON operation_work_logs(batch_id)"
    )
    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_operation_batches_date "
        "ON operation_batches(tarih, operator_ismi)"
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


def combination_name_from_operations(operations):
    """Operasyon etaplarının baş harflerinden okunabilir kombinasyon adı üretir."""
    initials = []

    for operation in operations:
        match = re.search(
            r"[A-Za-zÇĞİÖŞÜçğıöşü0-9]",
            str(operation),
        )
        if match:
            initials.append(match.group(0).upper())

    return "-".join(initials) or "KOMBİNASYON"


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
            base_name = combination_name_from_operations(ops)
            combination_name = base_name
            suffix = 2
            existing_names = {c["ad"] for c in combinations}

            while combination_name in existing_names:
                combination_name = f"{base_name} ({suffix})"
                suffix += 1

            combinations.append(
                {
                    "ad": combination_name,
                    "operasyonlar": ops,
                }
            )

    if not reasons:
        reasons = ["Yarı zamanlı çalışma", "Bakım", "İzinli", "Mazeret", "Hastalık", "Eğitim", "Boş", "Diğer"]

    if not combinations:
        default_operations = [
            "laser",
            "abkant",
            "çıta",
            "kaynak",
            "taşlama",
            "boya",
            "paketleme",
            "sevkiyat",
        ]
        combinations = [
            {
                "ad": combination_name_from_operations(default_operations),
                "operasyonlar": default_operations,
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
                COALESCE(i.combination_name, '') AS combination_name,
                i.source_name,
                i.imported_at,
                COALESCE(
                    SUM(
                        CASE
                            WHEN l.source_type = 'operation_tracking'
                            THEN l.produced_qty
                            ELSE 0
                        END
                    ),
                    0
                ) AS produced_qty
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
        deleted_operation_logs = 0
        if item_ids:
            placeholders = ",".join("?" for _ in item_ids)
            cur.execute(
                f"DELETE FROM operation_work_logs "
                f"WHERE item_id IN ({placeholders})",
                item_ids,
            )
            deleted_operation_logs = int(cur.rowcount or 0)
            cur.execute(
                f"DELETE FROM pos_operation_plan "
                f"WHERE item_id IN ({placeholders})",
                item_ids,
            )
            cur.execute(
                f"DELETE FROM production_output_logs "
                f"WHERE item_id IN ({placeholders})",
                item_ids,
            )
            deleted_logs = int(cur.rowcount or 0)

        cur.execute(
            "DELETE FROM operation_batches "
            "WHERE id NOT IN (SELECT DISTINCT batch_id FROM operation_work_logs)"
        )
        cur.execute(
            "DELETE FROM production_output_items WHERE oc_no = ?",
            (str(oc_no),),
        )
        deleted_items = int(cur.rowcount or 0)
        conn.commit()

        return {
            "deleted_items": deleted_items,
            "deleted_logs": deleted_logs,
            "deleted_operation_logs": deleted_operation_logs,
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


def _format_m2(value_mm2) -> str:
    numeric = float(value_mm2 or 0) / 1_000_000
    return f"{numeric:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")


def render_pos_area_overview(
    boy_mm: float,
    en_mm: float,
    unit_area_mm2: float,
    requested_qty: int,
    completed_qty: int,
):
    requested_area_mm2 = float(unit_area_mm2) * int(requested_qty)
    remaining_qty = max(int(requested_qty) - int(completed_qty), 0)
    completion_pct = (
        min(max(int(completed_qty) / int(requested_qty) * 100, 0), 100)
        if int(requested_qty) > 0
        else 0
    )
    ring_degrees = completion_pct * 3.6

    st.markdown(
        f"""
        <div class="area-stage">
            <div class="area-blueprint">
                <div class="blueprint-kicker">MİMARİ ÖLÇÜM PLANI</div>
                <div class="sheet-diagram">
                    <div class="dimension-x">{_format_mm2(boy_mm)} mm</div>
                    <div class="dimension-y">{_format_mm2(en_mm)} mm</div>
                </div>
                <div class="blueprint-main">{_format_m2(unit_area_mm2)} m²</div>
                <div class="blueprint-sub">
                    Birim alan · {_format_mm2(unit_area_mm2)} mm²
                </div>
            </div>
            <div class="area-stat-grid">
                <div class="area-stat-card">
                    <div class="area-stat-icon">◫</div>
                    <div class="area-stat-label">Parça ölçüsü</div>
                    <div class="area-stat-value">
                        {_format_mm2(boy_mm)} × {_format_mm2(en_mm)}
                    </div>
                    <div class="area-stat-detail">Boy × En · milimetre</div>
                </div>
                <div class="area-stat-card">
                    <div class="area-stat-icon">▦</div>
                    <div class="area-stat-label">İstenen toplam alan</div>
                    <div class="area-stat-value">
                        {_format_m2(requested_area_mm2)} m²
                    </div>
                    <div class="area-stat-detail">
                        {int(requested_qty)} adet üretim planı
                    </div>
                </div>
                <div class="area-stat-card completion-card">
                    <div
                        class="completion-ring"
                        style="--ring-value: {ring_degrees:.1f}deg;"
                    >
                        %{completion_pct:.0f}
                    </div>
                    <div>
                        <div class="area-stat-label">Tamamlanma</div>
                        <div class="area-stat-value">
                            {int(completed_qty)} / {int(requested_qty)}
                        </div>
                        <div class="area-stat-detail">
                            {remaining_qty} adet kaldı
                        </div>
                    </div>
                </div>
                <div class="area-stat-card">
                    <div class="area-stat-icon">⌁</div>
                    <div class="area-stat-label">Kalan alan</div>
                    <div class="area-stat-value">
                        {_format_m2(unit_area_mm2 * remaining_qty)} m²
                    </div>
                    <div class="area-stat-detail">
                        Kalan üretim ihtiyacı
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_area_summary(performance: dict, entries: list[dict]):
    processed_area_mm2 = sum(
        int(entry.get("processed_qty", 0))
        * float(entry.get("unit_area_mm2", 0))
        for entry in entries
    )
    fire_area_mm2 = sum(
        int(entry.get("fire_qty", 0))
        * float(entry.get("unit_area_mm2", 0))
        for entry in entries
    )
    good_area_mm2 = float(performance.get("total_good_area_mm2", 0))
    area_per_hour = float(performance.get("area_per_hour", 0))

    st.markdown(
        f"""
        <div class="area-summary">
            <div class="area-summary-item">
                <div class="area-summary-label">Sağlam üretim alanı</div>
                <div class="area-summary-value">{_format_m2(good_area_mm2)} m²</div>
                <div class="area-summary-detail">
                    Net ilerleyen · {_format_mm2(good_area_mm2)} mm²
                </div>
            </div>
            <div class="area-summary-item">
                <div class="area-summary-label">Alan verimi</div>
                <div class="area-summary-value">
                    {_format_m2(area_per_hour)} m²/saat
                </div>
                <div class="area-summary-detail">
                    Çalışma süresine göre anlık verim
                </div>
            </div>
            <div class="area-summary-item">
                <div class="area-summary-label">Fire alanı</div>
                <div class="area-summary-value">{_format_m2(fire_area_mm2)} m²</div>
                <div class="area-summary-detail">
                    İşlenen toplam alan: {_format_m2(processed_area_mm2)} m²
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )



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




def all_operations_from_combinations(combinations) -> list[str]:
    seen = set()
    result = []
    for combination in combinations:
        for operation in combination.get("operasyonlar", []):
            name = str(operation).strip()
            if name and name not in seen:
                seen.add(name)
                result.append(name)
    return result



def ensure_performance_targets(operations: list[str]):
    operation_names = []
    seen = set()
    for operation in operations:
        name = str(operation).strip()
        if name and name not in seen:
            seen.add(name)
            operation_names.append(name)

    if not operation_names:
        return

    conn = sqlite3.connect(DB_PATH)
    try:
        now = datetime.now().isoformat(timespec="seconds")
        conn.executemany(
            """
            INSERT OR IGNORE INTO operation_performance_targets (
                operation_name,
                target_qty_per_hour,
                target_area_per_hour,
                fire_limit_pct,
                slow_limit_pct,
                fast_limit_pct,
                updated_at
            ) VALUES (?, 0, 0, 5, 80, 120, ?)
            """,
            [(operation, now) for operation in operation_names],
        )
        conn.commit()
    finally:
        conn.close()


def get_performance_targets() -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql_query(
            """
            SELECT
                operation_name,
                target_qty_per_hour,
                target_area_per_hour,
                fire_limit_pct,
                slow_limit_pct,
                fast_limit_pct,
                updated_at
            FROM operation_performance_targets
            ORDER BY operation_name
            """,
            conn,
        )
    finally:
        conn.close()


def save_performance_targets(targets: pd.DataFrame):
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        now = datetime.now().isoformat(timespec="seconds")

        for _, row in targets.iterrows():
            operation_name = str(row["operation_name"]).strip()
            if not operation_name:
                continue

            qty_target = max(float(row["target_qty_per_hour"]), 0.0)
            area_target = max(float(row["target_area_per_hour"]), 0.0)
            fire_limit = min(max(float(row["fire_limit_pct"]), 0.0), 100.0)
            slow_limit = min(max(float(row["slow_limit_pct"]), 0.0), 500.0)
            fast_limit = min(max(float(row["fast_limit_pct"]), 0.0), 500.0)

            if fast_limit <= slow_limit:
                raise ValueError(
                    f"{operation_name}: Hızlı sınırı yavaş sınırından büyük olmalıdır."
                )

            cur.execute(
                """
                INSERT INTO operation_performance_targets (
                    operation_name,
                    target_qty_per_hour,
                    target_area_per_hour,
                    fire_limit_pct,
                    slow_limit_pct,
                    fast_limit_pct,
                    updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(operation_name) DO UPDATE SET
                    target_qty_per_hour = excluded.target_qty_per_hour,
                    target_area_per_hour = excluded.target_area_per_hour,
                    fire_limit_pct = excluded.fire_limit_pct,
                    slow_limit_pct = excluded.slow_limit_pct,
                    fast_limit_pct = excluded.fast_limit_pct,
                    updated_at = excluded.updated_at
                """,
                (
                    operation_name,
                    qty_target,
                    area_target,
                    fire_limit,
                    slow_limit,
                    fast_limit,
                    now,
                ),
            )

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_performance_target_map() -> dict[str, dict]:
    targets = get_performance_targets()
    if targets.empty:
        return {}

    return {
        str(row["operation_name"]): {
            "target_qty_per_hour": float(row["target_qty_per_hour"]),
            "target_area_per_hour": float(row["target_area_per_hour"]),
            "fire_limit_pct": float(row["fire_limit_pct"]),
            "slow_limit_pct": float(row["slow_limit_pct"]),
            "fast_limit_pct": float(row["fast_limit_pct"]),
        }
        for _, row in targets.iterrows()
    }


def performance_status(
    score_pct: float | None,
    slow_limit_pct: float = 80.0,
    fast_limit_pct: float = 120.0,
) -> str:
    if score_pct is None or pd.isna(score_pct):
        return "Referans yok"
    if float(score_pct) < float(slow_limit_pct):
        return "Yavaş"
    if float(score_pct) >= float(fast_limit_pct):
        return "Hızlı"
    return "Normal"


def get_historical_efficiency_reference() -> dict:
    conn = sqlite3.connect(DB_PATH)
    try:
        history = pd.read_sql_query(
            """
            SELECT
                b.id AS batch_id,
                b.calisma_saati,
                SUM(w.good_qty) AS good_qty,
                SUM(w.good_qty * i.unit_area_mm2) AS good_area_mm2
            FROM operation_batches b
            INNER JOIN operation_work_logs w ON w.batch_id = b.id
            INNER JOIN production_output_items i ON i.id = w.item_id
            WHERE b.calisma_saati > 0
            GROUP BY b.id, b.calisma_saati
            """,
            conn,
        )
    finally:
        conn.close()

    if history.empty:
        return {
            "qty_per_hour": 0.0,
            "area_per_hour": 0.0,
            "record_count": 0,
        }

    history["qty_per_hour"] = (
        history["good_qty"] / history["calisma_saati"].replace(0, 1)
    )
    history["area_per_hour"] = (
        history["good_area_mm2"] / history["calisma_saati"].replace(0, 1)
    )

    return {
        "qty_per_hour": float(history["qty_per_hour"].median()),
        "area_per_hour": float(history["area_per_hour"].median()),
        "record_count": int(len(history)),
    }


def calculate_entries_performance(
    entries: list[dict],
    work_hours: float,
) -> dict:
    total_processed = sum(int(entry.get("processed_qty", 0)) for entry in entries)
    total_fire = sum(int(entry.get("fire_qty", 0)) for entry in entries)
    total_good = sum(max(int(entry.get("good_qty", 0)), 0) for entry in entries)
    total_good_area = sum(
        max(int(entry.get("good_qty", 0)), 0)
        * float(entry.get("unit_area_mm2", 0))
        for entry in entries
    )

    safe_hours = float(work_hours) if float(work_hours) > 0 else 0.0
    qty_per_hour = total_good / safe_hours if safe_hours else 0.0
    area_per_hour = total_good_area / safe_hours if safe_hours else 0.0
    fire_rate = (
        total_fire / total_processed * 100
        if total_processed > 0
        else 0.0
    )

    target_map = get_performance_target_map()
    expected_qty_hours = 0.0
    expected_area_hours = 0.0
    qty_target_rows = 0
    area_target_rows = 0
    weighted_slow = []
    weighted_fast = []
    fire_breaches = []

    for entry in entries:
        operation_name = str(entry.get("operation_name", ""))
        target = target_map.get(
            operation_name,
            {
                "target_qty_per_hour": 0.0,
                "target_area_per_hour": 0.0,
                "fire_limit_pct": 5.0,
                "slow_limit_pct": 80.0,
                "fast_limit_pct": 120.0,
            },
        )

        good_qty = max(int(entry.get("good_qty", 0)), 0)
        processed_qty = max(int(entry.get("processed_qty", 0)), 0)
        fire_qty = max(int(entry.get("fire_qty", 0)), 0)
        area = good_qty * float(entry.get("unit_area_mm2", 0))

        if target["target_qty_per_hour"] > 0 and good_qty > 0:
            expected_qty_hours += (
                good_qty / target["target_qty_per_hour"]
            )
            qty_target_rows += 1

        if target["target_area_per_hour"] > 0 and area > 0:
            expected_area_hours += (
                area / target["target_area_per_hour"]
            )
            area_target_rows += 1

        weight = max(good_qty, 1)
        weighted_slow.extend([target["slow_limit_pct"]] * weight)
        weighted_fast.extend([target["fast_limit_pct"]] * weight)

        entry_fire_rate = (
            fire_qty / processed_qty * 100
            if processed_qty > 0
            else 0.0
        )
        if (
            processed_qty > 0
            and entry_fire_rate > target["fire_limit_pct"]
        ):
            fire_breaches.append(
                {
                    "pos": str(entry.get("pos", "")),
                    "operation": operation_name,
                    "fire_rate_pct": entry_fire_rate,
                    "fire_limit_pct": target["fire_limit_pct"],
                }
            )

    slow_limit = (
        sum(weighted_slow) / len(weighted_slow)
        if weighted_slow
        else 80.0
    )
    fast_limit = (
        sum(weighted_fast) / len(weighted_fast)
        if weighted_fast
        else 120.0
    )

    ratios = []
    reference_source = ""

    if safe_hours > 0 and qty_target_rows > 0:
        ratios.append(expected_qty_hours / safe_hours * 100)
    if safe_hours > 0 and area_target_rows > 0:
        ratios.append(expected_area_hours / safe_hours * 100)

    if ratios:
        score_pct = sum(ratios) / len(ratios)
        reference_source = "Yönetici hedefleri"
    else:
        historical = get_historical_efficiency_reference()
        historical_ratios = []
        if historical["record_count"] > 0:
            if historical["qty_per_hour"] > 0:
                historical_ratios.append(
                    qty_per_hour / historical["qty_per_hour"] * 100
                )
            if historical["area_per_hour"] > 0:
                historical_ratios.append(
                    area_per_hour / historical["area_per_hour"] * 100
                )

        if historical_ratios:
            score_pct = sum(historical_ratios) / len(historical_ratios)
            reference_source = "Geçmiş kayıt ortalaması"
        else:
            score_pct = None
            reference_source = "Henüz hedef veya geçmiş referans yok"

    return {
        "total_processed": total_processed,
        "total_fire": total_fire,
        "total_good": total_good,
        "total_good_area_mm2": total_good_area,
        "qty_per_hour": qty_per_hour,
        "area_per_hour": area_per_hour,
        "fire_rate_pct": fire_rate,
        "score_pct": score_pct,
        "status": performance_status(
            score_pct,
            slow_limit,
            fast_limit,
        ),
        "reference_source": reference_source,
        "slow_limit_pct": slow_limit,
        "fast_limit_pct": fast_limit,
        "fire_breaches": fire_breaches,
    }


def enrich_operation_batch_performance(
    batches: pd.DataFrame,
) -> pd.DataFrame:
    if batches.empty:
        return batches

    result = batches.copy()
    result["adet_saat"] = (
        result["saglam_ilerleyen"]
        / result["calisma_saati"].replace(0, 1)
    ).round(2)
    result["mm2_saat"] = (
        result["saglam_alan_mm2"]
        / result["calisma_saati"].replace(0, 1)
    ).round(2)
    result["fire_yuzde"] = (
        result["fire"]
        / result["islem_yapilan"].replace(0, 1)
        * 100
    ).round(2)

    scores = []
    statuses = []
    references = []

    historical = get_historical_efficiency_reference()

    for _, row in result.iterrows():
        ratios = []

        if float(row.get("target_qty_hours", 0) or 0) > 0:
            ratios.append(
                float(row["target_qty_hours"])
                / max(float(row["calisma_saati"]), 0.0001)
                * 100
            )
        if float(row.get("target_area_hours", 0) or 0) > 0:
            ratios.append(
                float(row["target_area_hours"])
                / max(float(row["calisma_saati"]), 0.0001)
                * 100
            )

        if ratios:
            score = sum(ratios) / len(ratios)
            reference = "Hedef"
        else:
            historical_ratios = []
            if historical["qty_per_hour"] > 0:
                historical_ratios.append(
                    float(row["adet_saat"])
                    / historical["qty_per_hour"]
                    * 100
                )
            if historical["area_per_hour"] > 0:
                historical_ratios.append(
                    float(row["mm2_saat"])
                    / historical["area_per_hour"]
                    * 100
                )
            score = (
                sum(historical_ratios) / len(historical_ratios)
                if historical_ratios
                else None
            )
            reference = (
                "Geçmiş ortalama"
                if historical_ratios
                else "Referans yok"
            )

        slow_limit = float(row.get("slow_limit_pct", 80) or 80)
        fast_limit = float(row.get("fast_limit_pct", 120) or 120)

        scores.append(round(score, 1) if score is not None else None)
        statuses.append(
            performance_status(score, slow_limit, fast_limit)
        )
        references.append(reference)

    result["verim_skoru_yuzde"] = scores
    result["gidisat"] = statuses
    result["verim_referansi"] = references
    result["fire_durumu"] = result.apply(
        lambda row: (
            "Sınır Aşıldı"
            if float(row["fire_yuzde"])
            > float(row.get("fire_limit_pct", 5) or 5)
            else "Normal"
        ),
        axis=1,
    )
    return result


def verim_fire_hedefleri_page(combinations):
    st.subheader("Verim ve Fire Hedefleri")
    st.caption(
        "Her işlem için adet/saat, mm²/saat ve fire sınırı belirle. "
        "Hedefi 0 bırakırsan gidişat geçmiş kayıt ortalamasına göre hesaplanır."
    )

    operations = all_operations_from_combinations(combinations)

    conn = sqlite3.connect(DB_PATH)
    try:
        plan_operations = pd.read_sql_query(
            """
            SELECT DISTINCT operation_name
            FROM pos_operation_plan
            ORDER BY operation_name
            """,
            conn,
        )
    finally:
        conn.close()

    if not plan_operations.empty:
        for operation in plan_operations["operation_name"].astype(str):
            if operation not in operations:
                operations.append(operation)

    ensure_performance_targets(operations)
    targets = get_performance_targets()

    if targets.empty:
        st.info("Henüz işlem bulunmuyor.")
        return

    editor = targets[
        [
            "operation_name",
            "target_qty_per_hour",
            "target_area_per_hour",
            "fire_limit_pct",
            "slow_limit_pct",
            "fast_limit_pct",
        ]
    ].rename(
        columns={
            "operation_name": "İşlem",
            "target_qty_per_hour": "Hedef Adet/Saat",
            "target_area_per_hour": "Hedef mm²/Saat",
            "fire_limit_pct": "Fire Sınırı %",
            "slow_limit_pct": "Yavaş Altı %",
            "fast_limit_pct": "Hızlı Üstü %",
        }
    )

    edited = st.data_editor(
        editor,
        use_container_width=True,
        hide_index=True,
        disabled=["İşlem"],
        column_config={
            "Hedef Adet/Saat": st.column_config.NumberColumn(
                min_value=0.0,
                step=0.1,
                format="%.2f",
            ),
            "Hedef mm²/Saat": st.column_config.NumberColumn(
                min_value=0.0,
                step=1000.0,
                format="%.2f",
            ),
            "Fire Sınırı %": st.column_config.NumberColumn(
                min_value=0.0,
                max_value=100.0,
                step=0.5,
                format="%.2f",
            ),
            "Yavaş Altı %": st.column_config.NumberColumn(
                min_value=0.0,
                max_value=500.0,
                step=5.0,
                format="%.1f",
            ),
            "Hızlı Üstü %": st.column_config.NumberColumn(
                min_value=0.0,
                max_value=500.0,
                step=5.0,
                format="%.1f",
            ),
        },
        key="performance_targets_editor",
    )

    if st.button(
        "Verim ve Fire Hedeflerini Kaydet",
        type="primary",
        use_container_width=True,
    ):
        save_frame = edited.rename(
            columns={
                "İşlem": "operation_name",
                "Hedef Adet/Saat": "target_qty_per_hour",
                "Hedef mm²/Saat": "target_area_per_hour",
                "Fire Sınırı %": "fire_limit_pct",
                "Yavaş Altı %": "slow_limit_pct",
                "Hızlı Üstü %": "fast_limit_pct",
            }
        )
        try:
            save_performance_targets(save_frame)
            st.success("Verim ve fire hedefleri kaydedildi.")
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))

    st.divider()
    st.markdown("### Gidişat yorumlama")
    st.write(
        "Verim skoru yavaş sınırının altındaysa **Yavaş**, hızlı sınırının "
        "üzerindeyse **Hızlı**, aradaysa **Normal** görünür."
    )
    st.write(
        "Fire oranı işlem için belirlediğin sınırı aşarsa yönetici ekranında "
        "**Fire sınırı aşıldı** uyarısı oluşur."
    )


def manager_fire_sidebar_alert():
    history = get_operation_history()
    if history.empty:
        return

    today_text = str(date.today())
    today = history[history["tarih"].astype(str) == today_text].copy()
    if today.empty:
        return

    target_map = get_performance_target_map()
    breaches = []

    for _, row in today.iterrows():
        processed = int(row["islem_yapilan"])
        fire = int(row["fire"])
        rate = fire / processed * 100 if processed else 0.0
        target = target_map.get(
            str(row["operasyon"]),
            {"fire_limit_pct": 5.0},
        )
        limit = float(target["fire_limit_pct"])
        if processed > 0 and rate > limit:
            breaches.append(
                f"{row['operator_ismi']} · {row['POS']} · "
                f"{row['operasyon']} %{rate:.1f}"
            )

    if breaches:
        st.error(f"Bugün {len(breaches)} fire sınırı uyarısı var.")
        with st.expander("Fire uyarılarını göster", expanded=False):
            for message in breaches:
                st.write("• " + message)



def get_worker_competencies(worker_name: str) -> list[str]:
    if not worker_name:
        return []
    conn = sqlite3.connect(DB_PATH)
    try:
        rows = conn.execute(
            "SELECT operation_name FROM worker_competencies "
            "WHERE worker_name = ? ORDER BY operation_name",
            (worker_name,),
        ).fetchall()
        return [str(row[0]) for row in rows]
    finally:
        conn.close()


def save_worker_competencies(worker_name: str, operations: list[str]):
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM worker_competencies WHERE worker_name = ?",
            (worker_name,),
        )
        created_at = datetime.now().isoformat(timespec="seconds")
        cur.executemany(
            "INSERT INTO worker_competencies "
            "(worker_name, operation_name, created_at) VALUES (?, ?, ?)",
            [(worker_name, operation, created_at) for operation in operations],
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_competency_table() -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql_query(
            """
            SELECT worker_name, operation_name
            FROM worker_competencies
            ORDER BY worker_name, operation_name
            """,
            conn,
        )
    finally:
        conn.close()


def get_item_operation_plan(item_id: int) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql_query(
            """
            SELECT combination_name, operation_order, operation_name
            FROM pos_operation_plan
            WHERE item_id = ?
            ORDER BY operation_order
            """,
            conn,
            params=(int(item_id),),
        )
    finally:
        conn.close()


def get_operation_progress(item_ids: list[int] | None = None) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        params = []
        item_filter = ""
        if item_ids:
            placeholders = ",".join("?" for _ in item_ids)
            item_filter = f" AND i.id IN ({placeholders})"
            params.extend(int(item_id) for item_id in item_ids)

        query = f"""
            WITH legacy AS (
                SELECT
                    item_id,
                    COALESCE(SUM(produced_qty), 0) AS legacy_qty
                FROM production_output_logs
                WHERE COALESCE(source_type, 'legacy') <> 'operation_tracking'
                GROUP BY item_id
            ),
            operation_totals AS (
                SELECT
                    item_id,
                    operation_name,
                    COALESCE(SUM(good_qty), 0) AS new_good_qty,
                    COALESCE(SUM(fire_qty), 0) AS fire_qty,
                    COALESCE(SUM(processed_qty), 0) AS processed_qty
                FROM operation_work_logs
                GROUP BY item_id, operation_name
            )
            SELECT
                i.id AS item_id,
                i.oc_no,
                COALESCE(NULLIF(i.project_name, ''), 'OC ' || i.oc_no) AS project_name,
                i.pos,
                i.requested_qty,
                i.boy_mm,
                i.en_mm,
                i.unit_area_mm2,
                p.combination_name,
                p.operation_order,
                p.operation_name,
                COALESCE(l.legacy_qty, 0) AS legacy_qty,
                COALESCE(o.new_good_qty, 0) AS operation_good_qty,
                COALESCE(o.fire_qty, 0) AS fire_qty,
                COALESCE(o.processed_qty, 0) AS processed_qty
            FROM pos_operation_plan p
            INNER JOIN production_output_items i ON i.id = p.item_id
            LEFT JOIN legacy l ON l.item_id = i.id
            LEFT JOIN operation_totals o
                ON o.item_id = i.id
                AND o.operation_name = p.operation_name
            WHERE 1 = 1 {item_filter}
            ORDER BY i.oc_no, i.pos, p.operation_order
        """
        result = pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()

    if result.empty:
        return result

    # Eski sürümlerdeki genel üretim kayıtları, belirli bir operasyonu
    # göstermediği için lazer/kaynak/boya gibi etaplara dağıtılamaz.
    # Operasyon ilerlemesi yalnızca operation_work_logs kayıtlarından hesaplanır.
    result["completed_qty"] = (
        result["operation_good_qty"]
    ).clip(upper=result["requested_qty"]).astype(int)
    result["remaining_qty"] = (
        result["requested_qty"] - result["completed_qty"]
    ).clip(lower=0).astype(int)
    result["completion_pct"] = (
        result["completed_qty"]
        / result["requested_qty"].replace(0, 1)
        * 100
    ).round(1)
    return result


def get_operation_history(oc_no: str | None = None, pos: str | None = None) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        query = """
            SELECT
                b.id AS batch_id,
                w.item_id AS item_id,
                b.tarih,
                b.operator_ismi,
                b.calisma_tipi,
                b.calisma_saati,
                i.oc_no AS OC,
                i.pos AS POS,
                w.operation_name AS operasyon,
                w.processed_qty AS islem_yapilan,
                w.fire_qty AS fire,
                w.good_qty AS saglam_ilerleyen,
                i.unit_area_mm2,
                w.good_qty * i.unit_area_mm2 AS saglam_alan_mm2,
                w.processed_qty * i.unit_area_mm2 AS islenen_alan_mm2,
                COALESCE(t.target_qty_per_hour, 0) AS hedef_adet_saat,
                COALESCE(t.target_area_per_hour, 0) AS hedef_mm2_saat,
                COALESCE(t.fire_limit_pct, 5) AS fire_siniri_yuzde,
                COALESCE(t.slow_limit_pct, 80) AS yavas_siniri_yuzde,
                COALESCE(t.fast_limit_pct, 120) AS hizli_siniri_yuzde,
                w.fire_note AS fire_notu,
                b.notlar,
                b.created_at
            FROM operation_work_logs w
            INNER JOIN operation_batches b ON b.id = w.batch_id
            INNER JOIN production_output_items i ON i.id = w.item_id
            LEFT JOIN operation_performance_targets t
                ON t.operation_name = w.operation_name
            WHERE 1 = 1
        """
        params = []
        if oc_no:
            query += " AND i.oc_no = ?"
            params.append(str(oc_no))
        if pos:
            query += " AND i.pos = ?"
            params.append(str(pos))
        query += " ORDER BY b.tarih DESC, b.id DESC, i.pos, w.operation_name"
        return pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()


def get_operation_overview() -> pd.DataFrame:
    output = get_production_output_summary()
    if output.empty:
        return output

    progress = get_operation_progress()
    overview = output.copy()
    overview["operation_fire_qty"] = 0
    overview["furthest_stage_qty"] = overview["produced_qty"]
    overview["planned_operations"] = 0

    if not progress.empty:
        grouped = progress.groupby("item_id", as_index=False).agg(
            operation_fire_qty=("fire_qty", "sum"),
            furthest_stage_qty=("completed_qty", "max"),
            planned_operations=("operation_name", "count"),
        )
        overview = overview.drop(
            columns=["operation_fire_qty", "furthest_stage_qty", "planned_operations"]
        ).merge(grouped, on="item_id", how="left")
        overview[["operation_fire_qty", "furthest_stage_qty", "planned_operations"]] = (
            overview[["operation_fire_qty", "furthest_stage_qty", "planned_operations"]]
            .fillna(0)
        )

    overview["in_production_qty"] = (
        overview["furthest_stage_qty"].astype(int)
        - overview["produced_qty"].astype(int)
    ).clip(lower=0)
    overview["completion_pct"] = (
        overview["produced_qty"]
        / overview["requested_qty"].replace(0, 1)
        * 100
    ).round(1)
    return overview


def save_operation_batch(batch_data: dict, entries: list[dict], plans: dict[int, dict]) -> int:
    if not entries:
        raise ValueError("Kaydedilecek işlem bulunamadı.")

    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        created_at = datetime.now().isoformat(timespec="seconds")
        cur.execute(
            """
            INSERT INTO operation_batches (
                tarih, operator_ismi, calisma_tipi, calisma_saati,
                neden, notlar, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                batch_data["tarih"],
                batch_data["operator_ismi"],
                batch_data["calisma_tipi"],
                float(batch_data["calisma_saati"]),
                batch_data.get("neden", ""),
                batch_data.get("notlar", ""),
                created_at,
            ),
        )
        batch_id = int(cur.lastrowid)

        affected_item_ids = sorted({int(entry["item_id"]) for entry in entries})
        before_completed = {}
        for item_id in affected_item_ids:
            cur.execute(
                "SELECT requested_qty, unit_area_mm2 FROM production_output_items WHERE id = ?",
                (item_id,),
            )
            item_row = cur.fetchone()
            if item_row is None:
                raise ValueError("POS kaydı bulunamadı.")
            cur.execute(
                """
                SELECT COALESCE(SUM(produced_qty), 0)
                FROM production_output_logs
                WHERE item_id = ?
                  AND source_type = 'operation_tracking'
                """,
                (item_id,),
            )
            before_completed[item_id] = int(cur.fetchone()[0] or 0)

        for item_id, plan in plans.items():
            cur.execute(
                "UPDATE production_output_items SET combination_name = ? WHERE id = ?",
                (plan["combination_name"], int(item_id)),
            )
            cur.execute("DELETE FROM pos_operation_plan WHERE item_id = ?", (int(item_id),))
            cur.executemany(
                """
                INSERT INTO pos_operation_plan (
                    item_id, combination_name, operation_order,
                    operation_name, created_at
                ) VALUES (?, ?, ?, ?, ?)
                """,
                [
                    (
                        int(item_id),
                        plan["combination_name"],
                        index,
                        operation,
                        created_at,
                    )
                    for index, operation in enumerate(plan["operations"], start=1)
                ],
            )

        for entry in entries:
            item_id = int(entry["item_id"])
            operation_name = str(entry["operation_name"])
            processed_qty = int(entry["processed_qty"])
            fire_qty = int(entry["fire_qty"])
            good_qty = processed_qty - fire_qty
            if processed_qty <= 0:
                continue
            if fire_qty < 0 or fire_qty > processed_qty:
                raise ValueError(f"{operation_name}: Fire, işlem yapılan adetten büyük olamaz.")

            cur.execute(
                "SELECT requested_qty FROM production_output_items WHERE id = ?",
                (item_id,),
            )
            requested_qty = int(cur.fetchone()[0])
            cur.execute(
                """
                SELECT COALESCE(SUM(good_qty), 0)
                FROM operation_work_logs
                WHERE item_id = ? AND operation_name = ?
                """,
                (item_id, operation_name),
            )
            existing_good = int(cur.fetchone()[0] or 0)
            if existing_good + good_qty > requested_qty:
                remaining = max(requested_qty - existing_good, 0)
                raise ValueError(
                    f"{operation_name}: Bu işlemde yalnızca {remaining} adet kaldı. "
                    f"Girilen sağlam adet {good_qty}."
                )

            cur.execute(
                """
                INSERT INTO operation_work_logs (
                    batch_id, item_id, operation_name, processed_qty,
                    fire_qty, good_qty, fire_note, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    batch_id,
                    item_id,
                    operation_name,
                    processed_qty,
                    fire_qty,
                    good_qty,
                    entry.get("fire_note", ""),
                    created_at,
                ),
            )

        for item_id in affected_item_ids:
            cur.execute(
                "SELECT requested_qty, unit_area_mm2 FROM production_output_items WHERE id = ?",
                (item_id,),
            )
            requested_qty, unit_area = cur.fetchone()
            cur.execute(
                """
                SELECT operation_name
                FROM pos_operation_plan
                WHERE item_id = ?
                ORDER BY operation_order
                """,
                (item_id,),
            )
            planned_operations = [str(row[0]) for row in cur.fetchall()]
            if not planned_operations:
                continue

            stage_totals = []
            for operation_name in planned_operations:
                cur.execute(
                    """
                    SELECT COALESCE(SUM(good_qty), 0)
                    FROM operation_work_logs
                    WHERE item_id = ? AND operation_name = ?
                    """,
                    (item_id, operation_name),
                )
                operation_good = int(cur.fetchone()[0] or 0)
                stage_totals.append(min(int(requested_qty), operation_good))

            after_completed = min(stage_totals) if stage_totals else 0
            delta = max(int(after_completed) - int(before_completed[item_id]), 0)
            if delta > 0:
                cur.execute(
                    """
                    INSERT INTO production_output_logs (
                        item_id, production_date, operator_name,
                        produced_qty, produced_area_mm2, note, created_at,
                        source_type, source_ref
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        item_id,
                        batch_data["tarih"],
                        batch_data["operator_ismi"],
                        delta,
                        delta * float(unit_area),
                        "Operasyonların tamamından geçen ürün",
                        created_at,
                        "operation_tracking",
                        batch_id,
                    ),
                )

        conn.commit()
        return batch_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_overtime_summary(target_date: str | None = None) -> pd.DataFrame:
    """Eski ve yeni kayıt sistemindeki mesaileri tek özet içinde gösterir."""
    conn = sqlite3.connect(DB_PATH)
    try:
        legacy_query = """
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
        params = []
        if target_date:
            legacy_query += " AND tarih = ?"
            params.append(str(target_date))
        legacy_query += " GROUP BY tarih, operator_ismi"
        legacy = pd.read_sql_query(legacy_query, conn, params=params)

        new_query = """
            SELECT
                b.tarih,
                b.operator_ismi,
                b.calisma_saati AS toplam_saat,
                (SELECT GROUP_CONCAT(DISTINCT i.oc_no)
                 FROM operation_work_logs w
                 INNER JOIN production_output_items i ON i.id = w.item_id
                 WHERE w.batch_id = b.id) AS OC,
                (SELECT GROUP_CONCAT(DISTINCT i.pos)
                 FROM operation_work_logs w
                 INNER JOIN production_output_items i ON i.id = w.item_id
                 WHERE w.batch_id = b.id) AS POS,
                (SELECT COALESCE(SUM(w.good_qty), 0)
                 FROM operation_work_logs w WHERE w.batch_id = b.id) AS uretilen_adet,
                (SELECT COALESCE(SUM(w.fire_qty), 0)
                 FROM operation_work_logs w WHERE w.batch_id = b.id) AS fire_adedi
            FROM operation_batches b
            WHERE b.calisma_tipi = 'Mesaili'
        """
        new_params = []
        if target_date:
            new_query += " AND b.tarih = ?"
            new_params.append(str(target_date))
        modern = pd.read_sql_query(new_query, conn, params=new_params)
    finally:
        conn.close()

    combined = pd.concat([legacy, modern], ignore_index=True)
    if combined.empty:
        return combined

    def join_unique(series):
        values = []
        for value in series.dropna().astype(str):
            for part in value.split(","):
                part = part.strip()
                if part and part not in values:
                    values.append(part)
        return ", ".join(values)

    result = combined.groupby(["tarih", "operator_ismi"], as_index=False).agg(
        toplam_saat=("toplam_saat", "sum"),
        OC=("OC", join_unique),
        POS=("POS", join_unique),
        uretilen_adet=("uretilen_adet", "sum"),
        fire_adedi=("fire_adedi", "sum"),
    )
    result["toplam_saat"] = result["toplam_saat"].round(2)
    return result.sort_values(["tarih", "operator_ismi"], ascending=[False, True])


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
                "uretilen_adet": "İşlem Adedi",
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
# DURLUM TASARIMI
# -----------------------------
def apply_durlum_theme():
    """Okunaklı, kurumsal ve Durlum ürünlerinden ilham alan arayüz."""
    st.markdown(
        """
        <style>
        :root {
            --df-navy-950: #062535;
            --df-navy-900: #0A3042;
            --df-navy-800: #12485F;
            --df-blue-600: #158AB5;
            --df-blue-500: #35ACCF;
            --df-cyan-300: #8AD9E4;
            --df-green-500: #48A996;
            --df-green-300: #89CFC0;
            --df-bg: #E8EEF1;
            --df-bg-soft: #F4F7F8;
            --df-surface: #FFFFFF;
            --df-surface-soft: #F5F8F9;
            --df-border: #CBD8DE;
            --df-border-soft: #DBE4E8;
            --df-text: #213A48;
            --df-muted: #667E8A;
            --df-success: #287B63;
            --df-success-bg: #E2F0EA;
            --df-warning: #A56A22;
            --df-warning-bg: #F7ECD9;
            --df-danger: #B6505D;
            --df-danger-bg: #F6E4E7;
            --df-purple: #6D66A3;
            --df-purple-bg: #ECE9F5;
            --df-shadow: 0 16px 42px rgba(8, 42, 58, 0.12);
            --df-shadow-soft: 0 7px 22px rgba(8, 42, 58, 0.075);
        }

        html, body, [class*="css"] {
            font-family: Inter, "Segoe UI", Arial, sans-serif;
        }

        .stApp {
            color: var(--df-text);
            background:
                radial-gradient(circle at 84% 5%, rgba(53, 172, 207, 0.13), transparent 25rem),
                radial-gradient(circle at 12% 95%, rgba(72, 169, 150, 0.09), transparent 30rem),
                linear-gradient(135deg, #E5ECEF 0%, #F4F7F8 50%, #E7EEF1 100%);
        }

        .stApp::before {
            content: "";
            position: fixed;
            inset: 0;
            pointer-events: none;
            z-index: 0;
            opacity: 0.24;
            background-image:
                radial-gradient(circle, rgba(10, 48, 66, 0.08) 0 1px, transparent 1.2px);
            background-size: 20px 20px;
            mask-image: linear-gradient(to bottom, rgba(0,0,0,.7), transparent 85%);
        }

        [data-testid="stAppViewContainer"] > .main {
            position: relative;
            z-index: 1;
        }

        .block-container {
            max-width: 1500px;
            padding-top: 0.9rem;
            padding-bottom: 3.5rem;
        }

        [data-testid="stHeader"] {
            background: rgba(244, 247, 248, 0.88);
            border-bottom: 1px solid rgba(159, 180, 190, 0.30);
            backdrop-filter: blur(14px);
        }

        [data-testid="stSidebar"] {
            color: #FFFFFF;
            background:
                radial-gradient(circle at 14% 0%, rgba(53, 172, 207, 0.22), transparent 14rem),
                radial-gradient(circle at 100% 24%, rgba(72, 169, 150, 0.15), transparent 12rem),
                linear-gradient(180deg, var(--df-navy-950), var(--df-navy-900));
            border-right: 1px solid rgba(138, 217, 228, 0.18);
            box-shadow: 14px 0 36px rgba(6, 37, 53, 0.18);
        }

        [data-testid="stSidebar"]::before {
            content: "";
            position: absolute;
            inset: 0;
            pointer-events: none;
            opacity: 0.13;
            background-image:
                radial-gradient(circle, rgba(255,255,255,.65) 0 1px, transparent 1.1px);
            background-size: 17px 17px;
        }

        [data-testid="stSidebar"] > div:first-child {
            padding-top: 0.75rem;
        }

        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] small,
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {
            color: #EAF7FA !important;
            opacity: 1 !important;
        }

        .sidebar-brand {
            position: relative;
            overflow: hidden;
            display: flex;
            align-items: center;
            gap: 0.76rem;
            padding: 0.82rem;
            margin-bottom: 0.75rem;
            border: 1px solid rgba(138, 217, 228, 0.21);
            border-radius: 16px;
            background: linear-gradient(
                145deg,
                rgba(255,255,255,.11),
                rgba(255,255,255,.035)
            );
            box-shadow: 0 13px 28px rgba(0,0,0,.14);
        }

        .sidebar-brand::after {
            content: "";
            position: absolute;
            width: 100px;
            height: 100px;
            right: -44px;
            top: -50px;
            border-radius: 50%;
            background: rgba(53, 172, 207, 0.18);
        }

        .sidebar-mark {
            position: relative;
            z-index: 1;
            width: 41px;
            height: 41px;
            display: grid;
            place-items: center;
            border-radius: 12px;
            color: #FFFFFF;
            background: linear-gradient(145deg, #3BB5D4, #107FA8 72%, #3B9F91);
            font-weight: 900;
            box-shadow: 0 11px 22px rgba(21, 138, 181, 0.28);
        }

        .sidebar-brand-title {
            position: relative;
            z-index: 1;
            color: #FFFFFF;
            font-size: 0.95rem;
            font-weight: 850;
        }

        .sidebar-brand-subtitle {
            position: relative;
            z-index: 1;
            margin-top: 0.13rem;
            color: #B9D7DF;
            font-size: 0.71rem;
        }

        .sidebar-version {
            padding: 0.54rem 0.7rem;
            margin-bottom: 0.75rem;
            border-radius: 11px;
            color: #D3F3F7;
            background: linear-gradient(
                145deg,
                rgba(53,172,207,.14),
                rgba(72,169,150,.08)
            );
            border: 1px solid rgba(138,217,228,.16);
            font-size: 0.75rem;
            font-weight: 750;
            text-align: center;
        }

        [data-testid="stSidebar"] .stRadio label {
            min-height: 2.48rem;
            display: flex;
            align-items: center;
            padding: 0.42rem 0.55rem;
            border-radius: 10px;
            border: 1px solid transparent;
            transition: background .14s ease, transform .14s ease;
        }

        [data-testid="stSidebar"] .stRadio label p,
        [data-testid="stSidebar"] .stRadio label span {
            color: #DDEEF3 !important;
            font-weight: 680 !important;
        }

        [data-testid="stSidebar"] .stRadio label:hover {
            background: rgba(138,217,228,.08);
            transform: translateX(2px);
        }

        [data-testid="stSidebar"] .stRadio label:hover p,
        [data-testid="stSidebar"] .stRadio label:hover span {
            color: #FFFFFF !important;
        }

        [data-testid="stSidebar"] .stRadio label:has(input:checked) {
            background: linear-gradient(
                90deg,
                rgba(53,172,207,.20),
                rgba(72,169,150,.07)
            );
            border-color: rgba(138,217,228,.18);
            box-shadow: inset 3px 0 0 var(--df-cyan-300);
        }

        [data-testid="stSidebar"] .stRadio label:has(input:checked) p,
        [data-testid="stSidebar"] .stRadio label:has(input:checked) span {
            color: #FFFFFF !important;
            font-weight: 820 !important;
        }

        [data-testid="stSidebar"] [data-testid="stExpander"] {
            background: rgba(255,255,255,.055) !important;
            border-color: rgba(138,217,228,.18) !important;
            box-shadow: none !important;
        }

        [data-testid="stSidebar"] [data-testid="stExpander"] summary p,
        [data-testid="stSidebar"] [data-testid="stExpander"] summary span,
        [data-testid="stSidebar"] [data-testid="stExpander"] summary svg {
            color: #FFFFFF !important;
            fill: #FFFFFF !important;
            opacity: 1 !important;
            font-weight: 800 !important;
        }

        [data-testid="stSidebar"] .notification-card {
            background: linear-gradient(
                145deg,
                rgba(255,255,255,.12),
                rgba(138,217,228,.045)
            ) !important;
            border-color: rgba(186,226,235,.22) !important;
            box-shadow: none !important;
        }

        [data-testid="stSidebar"] .notification-title,
        [data-testid="stSidebar"] .notification-title * {
            color: #FFFFFF !important;
        }

        [data-testid="stSidebar"] .notification-copy,
        [data-testid="stSidebar"] .notification-copy * {
            color: #D1E8EE !important;
        }

        .durlum-header {
            position: relative;
            overflow: hidden;
            display: grid;
            grid-template-columns: minmax(0, 1.45fr) minmax(330px, .72fr);
            min-height: 205px;
            margin: 0.05rem 0 0.9rem;
            border: 1px solid rgba(112,149,164,.34);
            border-radius: 26px;
            background: linear-gradient(
                135deg,
                #062535 0%,
                #0D354B 54%,
                #0A2D40 100%
            );
            box-shadow: var(--df-shadow);
        }

        .durlum-header::before {
            content: "";
            position: absolute;
            inset: 0;
            opacity: 0.17;
            background-image:
                radial-gradient(circle, rgba(255,255,255,.88) 0 1px, transparent 1.15px);
            background-size: 18px 18px;
            mask-image: linear-gradient(90deg, transparent 0 46%, #000 46%);
        }

        .durlum-header::after {
            content: "";
            position: absolute;
            inset: auto 0 0;
            height: 4px;
            background: linear-gradient(
                90deg,
                transparent,
                var(--df-blue-500) 20%,
                var(--df-cyan-300) 48%,
                var(--df-green-300) 70%,
                transparent
            );
        }

        .durlum-hero-copy {
            position: relative;
            z-index: 2;
            display: flex;
            gap: 1rem;
            padding: 1.25rem 1.35rem;
        }

        .durlum-logo-box {
            width: 68px;
            height: 68px;
            min-width: 68px;
            display: grid;
            place-items: center;
            border-radius: 18px;
            color: #FFFFFF;
            background:
                linear-gradient(145deg, rgba(255,255,255,.14), rgba(255,255,255,.02)),
                linear-gradient(145deg, #38B2D2, #107FA8 68%, #3CA393);
            border: 1px solid rgba(255,255,255,.14);
            box-shadow: 0 14px 28px rgba(0,0,0,.18);
            font-size: 1rem;
            font-weight: 900;
            line-height: .9;
            text-align: center;
        }

        .durlum-eyebrow {
            display: inline-flex;
            align-items: center;
            gap: 0.42rem;
            padding: 0.34rem 0.58rem;
            border-radius: 999px;
            color: #D0F6FA;
            background: rgba(255,255,255,.07);
            border: 1px solid rgba(255,255,255,.11);
            font-size: 0.68rem;
            font-weight: 790;
            letter-spacing: .13em;
            text-transform: uppercase;
        }

        .durlum-eyebrow-dot {
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background: linear-gradient(145deg, #9CE8F0, #43B8D4);
            box-shadow: 0 0 0 4px rgba(138,217,228,.10);
        }

        .product-ribbon {
            display: inline-flex;
            margin-top: 0.74rem;
            padding: 0.4rem 0.62rem;
            border-radius: 11px;
            color: #D8F4F8;
            background: rgba(255,255,255,.065);
            border: 1px solid rgba(255,255,255,.10);
            font-size: 0.73rem;
            font-weight: 760;
        }

        .durlum-title {
            margin: 0.8rem 0 0.32rem;
            color: #F7FEFF;
            font-size: clamp(1.75rem, 3vw, 2.55rem);
            line-height: 1;
            letter-spacing: -.048em;
        }

        .durlum-title span {
            color: var(--df-cyan-300);
        }

        .durlum-subtitle {
            max-width: 720px;
            color: #C3DCE4;
            font-size: 0.92rem;
            line-height: 1.56;
        }

        .durlum-header-side {
            display: flex;
            flex-wrap: wrap;
            gap: 0.55rem;
            margin-top: 0.82rem;
        }

        .durlum-role,
        .durlum-date {
            display: inline-flex;
            align-items: center;
            padding: 0.45rem 0.68rem;
            border-radius: 11px;
            color: #ECFBFD;
            background: rgba(255,255,255,.07);
            border: 1px solid rgba(255,255,255,.10);
            font-size: 0.76rem;
            font-weight: 740;
        }

        .hero-art {
            position: relative;
            z-index: 1;
            min-height: 205px;
            overflow: hidden;
            background:
                linear-gradient(135deg, rgba(255,255,255,.015), rgba(255,255,255,.055));
        }

        .mesh-panel {
            position: absolute;
            border: 1px solid rgba(190,239,245,.22);
            border-radius: 20px;
            background:
                radial-gradient(circle, rgba(219,248,251,.46) 0 1.4px, transparent 1.6px),
                linear-gradient(145deg, rgba(255,255,255,.055), rgba(138,217,228,.10));
            background-size: 15px 15px, auto;
            box-shadow: inset 0 0 0 1px rgba(255,255,255,.03);
        }

        .mesh-panel.panel-one {
            width: 185px;
            height: 92px;
            right: 160px;
            top: 27px;
            transform: rotate(-8deg);
        }

        .mesh-panel.panel-two {
            width: 205px;
            height: 105px;
            right: 38px;
            top: 84px;
            transform: rotate(7deg);
            background-size: 13px 13px, auto;
        }

        .mesh-panel.panel-three {
            width: 145px;
            height: 72px;
            right: 24px;
            top: 22px;
            transform: rotate(-3deg);
            background-size: 18px 18px, auto;
        }

        .ceiling-lines {
            position: absolute;
            left: 20px;
            right: 10px;
            bottom: 26px;
            height: 72px;
            opacity: 0.35;
            background:
                repeating-linear-gradient(
                    -24deg,
                    transparent 0 24px,
                    rgba(138,217,228,.48) 24px 26px,
                    transparent 26px 52px
                );
        }

        .page-spotlight {
            position: relative;
            overflow: hidden;
            display: flex;
            align-items: center;
            gap: 0.8rem;
            padding: 0.85rem 0.95rem;
            margin: 0.28rem 0 1rem;
            border: 1px solid rgba(157,180,190,.38);
            border-radius: 17px;
            background: linear-gradient(
                145deg,
                rgba(255,255,255,.90),
                rgba(241,246,247,.94)
            );
            box-shadow: var(--df-shadow-soft);
        }

        .page-spotlight::after {
            content: "";
            position: absolute;
            right: -36px;
            top: -48px;
            width: 150px;
            height: 150px;
            border-radius: 50%;
            background: rgba(72,169,150,.08);
        }

        .page-spotlight-icon {
            position: relative;
            z-index: 1;
            width: 43px;
            height: 43px;
            min-width: 43px;
            display: grid;
            place-items: center;
            border-radius: 12px;
            color: var(--df-blue-600);
            background: linear-gradient(145deg, #EAF5F7, #DDEEF3);
            border: 1px solid #BED7E0;
            font-size: 1.13rem;
        }

        .page-spotlight-title {
            position: relative;
            z-index: 1;
            color: var(--df-text);
            font-size: 0.97rem;
            font-weight: 840;
        }

        .page-spotlight-copy {
            position: relative;
            z-index: 1;
            margin-top: 0.15rem;
            color: var(--df-muted);
            font-size: 0.79rem;
        }

        h1, h2, h3 {
            color: var(--df-text);
            letter-spacing: -.027em;
        }

        [data-testid="stMetric"] {
            position: relative;
            overflow: hidden;
            min-height: 112px;
            padding: 0.94rem 1rem;
            border: 1px solid var(--df-border);
            border-radius: 17px;
            background: linear-gradient(
                145deg,
                rgba(255,255,255,.96),
                rgba(239,245,247,.96)
            );
            box-shadow: var(--df-shadow-soft);
        }

        [data-testid="stMetric"]::before {
            content: "";
            position: absolute;
            inset: 0 auto 0 0;
            width: 4px;
            background: linear-gradient(
                180deg,
                var(--df-blue-500),
                var(--df-green-500)
            );
        }

        [data-testid="stMetric"]::after {
            content: "";
            position: absolute;
            width: 105px;
            height: 105px;
            right: -46px;
            top: -54px;
            border-radius: 50%;
            background: radial-gradient(
                circle,
                rgba(53,172,207,.14),
                transparent 68%
            );
        }

        [data-testid="stMetricLabel"] p {
            color: var(--df-muted) !important;
            font-size: 0.73rem !important;
            font-weight: 750 !important;
        }

        [data-testid="stMetricValue"] {
            color: var(--df-text) !important;
            font-weight: 880 !important;
            letter-spacing: -.04em;
        }

        .wizard-shell {
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 0.42rem;
            margin: 0.28rem 0 1rem;
            padding: 0.52rem;
            border: 1px solid var(--df-border);
            border-radius: 17px;
            background: linear-gradient(
                145deg,
                rgba(232,239,242,.95),
                rgba(245,248,249,.96)
            );
            box-shadow: var(--df-shadow-soft);
        }

        .wizard-step {
            min-height: 65px;
            padding: 0.58rem;
            border-radius: 11px;
            border: 1px solid transparent;
            color: #788D98;
            background: rgba(255,255,255,.54);
        }

        .wizard-step-number {
            display: inline-grid;
            place-items: center;
            width: 23px;
            height: 23px;
            margin-bottom: 0.28rem;
            border-radius: 50%;
            color: #6E838D;
            background: #D8E3E7;
            font-size: 0.7rem;
            font-weight: 900;
        }

        .wizard-step-title {
            display: block;
            font-size: 0.75rem;
            font-weight: 820;
        }

        .wizard-step.done {
            color: var(--df-success);
            background: var(--df-success-bg);
            border-color: #C2DBD0;
        }

        .wizard-step.done .wizard-step-number {
            color: #FFFFFF;
            background: linear-gradient(145deg, #31977A, #287B63);
        }

        .wizard-step.active {
            color: var(--df-blue-600);
            background: #FFFFFF;
            border-color: #A7CBD7;
            box-shadow:
                inset 0 -3px 0 var(--df-blue-600),
                0 8px 20px rgba(8,42,58,.07);
        }

        .wizard-step.active .wizard-step-number {
            color: #FFFFFF;
            background: linear-gradient(145deg, #35ACCF, #158AB5);
        }

        .wizard-panel,
        .review-card,
        .insight-card {
            border: 1px solid var(--df-border);
            background: linear-gradient(
                145deg,
                rgba(255,255,255,.96),
                rgba(240,245,247,.96)
            );
            box-shadow: var(--df-shadow-soft);
        }

        .wizard-panel {
            position: relative;
            overflow: hidden;
            padding: 0.98rem 1rem;
            margin-bottom: 0.88rem;
            border-radius: 17px;
        }

        .wizard-panel-title,
        .review-title,
        .insight-title {
            color: var(--df-text);
            font-weight: 840;
        }

        .wizard-panel-copy,
        .review-label,
        .insight-copy {
            color: var(--df-muted);
        }

        .review-card {
            padding: 0.94rem 1rem;
            margin: 0.4rem 0 0.9rem;
            border-radius: 17px;
        }

        .review-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 0.56rem;
        }

        .review-item {
            padding: 0.64rem 0.7rem;
            border: 1px solid var(--df-border-soft);
            border-radius: 11px;
            background: linear-gradient(
                145deg,
                rgba(232,240,243,.80),
                rgba(251,252,252,.94)
            );
        }

        .review-value {
            margin-top: 0.16rem;
            color: var(--df-text);
            font-size: 0.92rem;
            font-weight: 840;
        }

        .status-pill {
            display: inline-flex;
            align-items: center;
            padding: 0.38rem 0.64rem;
            border-radius: 999px;
            border: 1px solid transparent;
            font-size: 0.75rem;
            font-weight: 820;
        }

        .status-pill.blue {
            color: #146D8C;
            background: #E0F0F4;
            border-color: #C3DCE5;
        }

        .status-pill.green {
            color: var(--df-success);
            background: var(--df-success-bg);
            border-color: #C1D9CF;
        }

        .status-pill.orange {
            color: var(--df-warning);
            background: var(--df-warning-bg);
            border-color: #E2CBA7;
        }

        .status-pill.red {
            color: var(--df-danger);
            background: var(--df-danger-bg);
            border-color: #E5C5CA;
        }

        .status-pill.purple {
            color: var(--df-purple);
            background: var(--df-purple-bg);
            border-color: #D7D1E8;
        }

        .status-pill.gray {
            color: #667A85;
            background: #E9EFF1;
            border-color: #D4DEE2;
        }

        .notification-card {
            position: relative;
            overflow: hidden;
            padding: 0.7rem 0.76rem;
            margin: 0.36rem 0;
            border: 1px solid var(--df-border);
            border-radius: 13px;
            background: linear-gradient(
                145deg,
                rgba(241,246,247,.95),
                rgba(252,253,253,.97)
            );
            box-shadow: var(--df-shadow-soft);
        }

        .notification-card::before {
            content: "";
            position: absolute;
            inset: 0 auto 0 0;
            width: 4px;
            background: linear-gradient(
                180deg,
                var(--df-blue-500),
                var(--df-green-500)
            );
        }

        .notification-title {
            color: var(--df-text);
            font-size: 0.77rem;
            font-weight: 830;
        }

        .notification-copy {
            margin-top: 0.18rem;
            color: var(--df-muted);
            font-size: 0.7rem;
            line-height: 1.38;
        }

        .empty-state {
            display: grid;
            place-items: center;
            min-height: 230px;
            padding: 1.35rem;
            text-align: center;
            border: 1px dashed #B5C9D1;
            border-radius: 19px;
            background:
                radial-gradient(circle at 50% 0%, rgba(53,172,207,.10), transparent 13rem),
                linear-gradient(145deg, #FAFCFC, #F0F6F7);
        }

        .empty-icon {
            width: 59px;
            height: 59px;
            display: grid;
            place-items: center;
            margin: 0 auto 0.7rem;
            border-radius: 15px;
            color: var(--df-blue-600);
            background: linear-gradient(145deg, #E6F3F6, #DCEFF3);
            font-size: 1.48rem;
        }

        .empty-title {
            color: var(--df-text);
            font-size: 1.05rem;
            font-weight: 840;
        }

        .empty-copy {
            max-width: 610px;
            margin-top: 0.32rem;
            color: var(--df-muted);
            font-size: 0.84rem;
            line-height: 1.48;
        }

        .completed-state {
            padding: 0.94rem 1rem;
            border-radius: 15px;
            color: #286F59;
            background: linear-gradient(145deg, #E4F2EC, #EDF8F3);
            border: 1px solid #C0D9CE;
            box-shadow: var(--df-shadow-soft);
        }

        .insight-card {
            position: relative;
            overflow: hidden;
            min-height: 115px;
            padding: 0.84rem 0.88rem;
            border-radius: 16px;
        }

        .insight-card::after {
            content: "";
            position: absolute;
            width: 100px;
            height: 100px;
            right: -42px;
            top: -50px;
            border-radius: 50%;
            background: radial-gradient(
                circle,
                rgba(53,172,207,.12),
                transparent 70%
            );
        }

        .insight-icon {
            position: relative;
            z-index: 1;
            color: var(--df-blue-600);
            font-size: 1.06rem;
        }

        .insight-title {
            position: relative;
            z-index: 1;
            margin-top: 0.38rem;
            font-size: 0.94rem;
        }

        .insight-copy {
            position: relative;
            z-index: 1;
            margin-top: 0.22rem;
            font-size: 0.72rem;
            line-height: 1.43;
        }

        .product-footer {
            margin-top: 1.5rem;
            padding: 0.86rem 1rem;
            border: 1px solid var(--df-border);
            border-radius: 15px;
            color: var(--df-muted);
            background: linear-gradient(
                145deg,
                rgba(255,255,255,.80),
                rgba(240,245,247,.90)
            );
            text-align: center;
            font-size: 0.75rem;
            box-shadow: var(--df-shadow-soft);
        }

        .stButton > button,
        [data-testid="baseButton-secondary"],
        [data-testid="baseButton-primary"] {
            min-height: 2.85rem;
            border-radius: 12px !important;
            border: 1px solid #BDD1D9 !important;
            color: var(--df-text) !important;
            background: linear-gradient(145deg, #FFFFFF, #EFF5F6) !important;
            box-shadow: 0 7px 18px rgba(8,42,58,.06);
            font-weight: 780 !important;
        }

        [data-testid="baseButton-primary"] {
            color: #FFFFFF !important;
            background: linear-gradient(
                145deg,
                #35ACCF,
                #158AB5 70%,
                #3E9F91
            ) !important;
            border-color: #158AB5 !important;
        }

        .stTextInput > div > div,
        .stNumberInput > div > div,
        .stDateInput > div > div,
        .stSelectbox > div > div,
        .stMultiSelect > div > div,
        .stTextArea textarea {
            border-radius: 12px !important;
            border: 1px solid #C6D5DB !important;
            background: rgba(255,255,255,.94) !important;
        }

        .stDataFrame,
        [data-testid="stTable"] {
            overflow: hidden;
            border: 1px solid var(--df-border);
            border-radius: 17px;
            background: rgba(255,255,255,.86);
            box-shadow: var(--df-shadow-soft);
        }

        [data-testid="stDataFrame"] [role="columnheader"] {
            background: linear-gradient(145deg, #E8EFF2, #F8FAFB) !important;
        }

        .area-blueprint {
            position: relative;
            overflow: hidden;
            min-height: 215px;
            padding: 1rem;
            border: 1px solid rgba(170,197,207,.96);
            border-radius: 19px;
            color: #EAFBFD;
            background:
                radial-gradient(circle, rgba(255,255,255,.13) 0 1px, transparent 1.2px),
                linear-gradient(rgba(138,217,228,.06) 1px, transparent 1px),
                linear-gradient(90deg, rgba(138,217,228,.06) 1px, transparent 1px),
                linear-gradient(145deg, #082B3D, #0E3950 68%, #0A3042);
            background-size: 16px 16px, 24px 24px, 24px 24px, auto;
            box-shadow: var(--df-shadow);
        }

        .area-stage {
            display: grid;
            grid-template-columns: minmax(310px, .95fr) minmax(0, 1.12fr);
            gap: 0.84rem;
            align-items: stretch;
            margin: 0.54rem 0 0.84rem;
        }

        .area-stat-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.66rem;
        }

        .area-stat-card {
            position: relative;
            overflow: hidden;
            min-height: 106px;
            padding: 0.8rem 0.84rem;
            border: 1px solid var(--df-border);
            border-radius: 15px;
            background: linear-gradient(
                145deg,
                rgba(255,255,255,.97),
                rgba(239,245,247,.96)
            );
            box-shadow: var(--df-shadow-soft);
        }

        .area-stat-card::before {
            content: "";
            position: absolute;
            inset: 0 auto 0 0;
            width: 4px;
            background: linear-gradient(
                180deg,
                var(--df-blue-500),
                var(--df-green-500)
            );
        }

        .area-stat-label,
        .area-summary-label {
            color: var(--df-muted);
            font-size: 0.69rem;
            font-weight: 760;
        }

        .area-stat-value {
            margin-top: 0.28rem;
            color: var(--df-text);
            font-size: 1.34rem;
            font-weight: 900;
            letter-spacing: -.04em;
        }

        .area-stat-detail,
        .area-summary-detail {
            margin-top: 0.14rem;
            color: var(--df-muted);
            font-size: 0.68rem;
            line-height: 1.4;
        }

        .area-summary {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 0.64rem;
            margin: 0.74rem 0 0.88rem;
        }

        .area-summary-item {
            position: relative;
            overflow: hidden;
            min-height: 108px;
            padding: 0.8rem 0.86rem;
            border: 1px solid var(--df-border);
            border-radius: 15px;
            background: linear-gradient(145deg, #E7F0F3, #FAFCFC);
            box-shadow: var(--df-shadow-soft);
        }

        .area-summary-item::before {
            content: "";
            position: absolute;
            inset: 0 auto 0 0;
            width: 4px;
            background: linear-gradient(
                180deg,
                var(--df-blue-500),
                var(--df-green-500)
            );
        }

        .area-summary-value {
            margin-top: 0.27rem;
            color: var(--df-blue-600);
            font-size: 1.2rem;
            font-weight: 880;
        }

        @media (max-width: 980px) {
            .durlum-header {
                grid-template-columns: 1fr;
            }

            .hero-art {
                display: none;
            }

            .area-stage {
                grid-template-columns: 1fr;
            }
        }

        @media (max-width: 900px) {
            .durlum-title {
                font-size: 2rem;
            }

            .review-grid,
            .wizard-shell {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }

            .area-summary {
                grid-template-columns: 1fr;
            }
        }

        @media (max-width: 620px) {
            .review-grid,
            .wizard-shell,
            .area-stat-grid {
                grid-template-columns: 1fr;
            }
        }

        /* SIDEBAR BUTON VE FORM OKUNABİLİRLİĞİ */
        [data-testid="stSidebar"] .stButton > button,
        [data-testid="stSidebar"] [data-testid="baseButton-secondary"],
        [data-testid="stSidebar"] [data-testid="baseButton-primary"] {
            min-height: 2.8rem !important;
            color: #FFFFFF !important;
            background:
                linear-gradient(
                    135deg,
                    #168FB8 0%,
                    #0E789F 68%,
                    #348F84 100%
                ) !important;
            border: 1px solid rgba(138, 217, 228, 0.34) !important;
            border-radius: 12px !important;
            box-shadow:
                0 9px 20px rgba(0, 0, 0, 0.16),
                inset 0 1px 0 rgba(255,255,255,.14) !important;
            font-weight: 820 !important;
        }

        [data-testid="stSidebar"] .stButton > button p,
        [data-testid="stSidebar"] .stButton > button span,
        [data-testid="stSidebar"] [data-testid="baseButton-secondary"] p,
        [data-testid="stSidebar"] [data-testid="baseButton-secondary"] span,
        [data-testid="stSidebar"] [data-testid="baseButton-primary"] p,
        [data-testid="stSidebar"] [data-testid="baseButton-primary"] span {
            color: #FFFFFF !important;
            opacity: 1 !important;
            font-weight: 820 !important;
        }

        [data-testid="stSidebar"] .stButton > button:hover,
        [data-testid="stSidebar"] [data-testid="baseButton-secondary"]:hover,
        [data-testid="stSidebar"] [data-testid="baseButton-primary"]:hover {
            color: #FFFFFF !important;
            background:
                linear-gradient(
                    135deg,
                    #25A9CF 0%,
                    #1185AC 65%,
                    #43A395 100%
                ) !important;
            border-color: rgba(177, 235, 242, 0.55) !important;
            transform: translateY(-1px);
        }

        [data-testid="stSidebar"] .stButton > button:disabled,
        [data-testid="stSidebar"] [data-testid="baseButton-secondary"]:disabled,
        [data-testid="stSidebar"] [data-testid="baseButton-primary"]:disabled {
            color: #AFC3CB !important;
            background: rgba(163, 188, 198, 0.14) !important;
            border-color: rgba(185, 211, 220, 0.18) !important;
            box-shadow: none !important;
        }

        [data-testid="stSidebar"] .stTextInput > div > div,
        [data-testid="stSidebar"] .stNumberInput > div > div,
        [data-testid="stSidebar"] .stDateInput > div > div,
        [data-testid="stSidebar"] .stSelectbox > div > div,
        [data-testid="stSidebar"] .stMultiSelect > div > div,
        [data-testid="stSidebar"] .stTextArea textarea {
            color: #183746 !important;
            background: #F7FAFB !important;
            border: 1px solid rgba(181, 211, 220, 0.72) !important;
            box-shadow:
                inset 0 1px 0 rgba(255,255,255,.75),
                0 5px 14px rgba(0,0,0,.08) !important;
        }

        [data-testid="stSidebar"] .stTextInput input,
        [data-testid="stSidebar"] .stNumberInput input,
        [data-testid="stSidebar"] .stDateInput input,
        [data-testid="stSidebar"] .stTextArea textarea,
        [data-testid="stSidebar"] div[data-baseweb="select"] input,
        [data-testid="stSidebar"] div[data-baseweb="select"] span {
            color: #183746 !important;
            -webkit-text-fill-color: #183746 !important;
            opacity: 1 !important;
        }

        [data-testid="stSidebar"] .stTextInput input::placeholder,
        [data-testid="stSidebar"] .stTextArea textarea::placeholder {
            color: #728B97 !important;
            -webkit-text-fill-color: #728B97 !important;
            opacity: 1 !important;
        }

        [data-testid="stSidebar"] .stTextInput svg,
        [data-testid="stSidebar"] .stNumberInput svg,
        [data-testid="stSidebar"] .stDateInput svg,
        [data-testid="stSidebar"] .stSelectbox svg,
        [data-testid="stSidebar"] .stMultiSelect svg {
            color: #476C7D !important;
            fill: #476C7D !important;
        }

        </style>
        """,
        unsafe_allow_html=True,
    )


def render_durlum_header(is_manager: bool):
    role = "Yönetici görünümü" if is_manager else "Çalışan görünümü"
    role_icon = "🛡️" if is_manager else "👷"
    today_text = date.today().strftime("%d.%m.%Y")

    header_html = (
        '<div class="durlum-header">'
        '<div class="durlum-hero-copy">'
        '<div class="durlum-logo-box">dur<br>lum</div>'
        '<div>'
        '<div class="durlum-eyebrow">'
        '<span class="durlum-eyebrow-dot"></span>'
        'ARCHITECTURAL PRODUCTION SYSTEM'
        '</div>'
        '<div class="product-ribbon">'
        'DURLUM FLOW · Mesh · Ceiling · Metal · Intelligence'
        '</div>'
        '<h1 class="durlum-title">'
        'Üretim <span>Akışı</span> ve Operasyon Merkezi'
        '</h1>'
        '<div class="durlum-subtitle">'
        'Mesh tavan, metal yüzey ve modüler üretim mantığından ilham alan '
        'dijital kontrol sistemi; OC, POS, operasyon, verim ve fire süreçlerini '
        'tek bir kurumsal arayüzde birleştirir.'
        '</div>'
        '<div class="durlum-header-side">'
        f'<div class="durlum-role">{role_icon} {role}</div>'
        f'<div class="durlum-date">📅 {today_text} · Sistem aktif</div>'
        '</div>'
        '</div>'
        '</div>'
        '<div class="hero-art" aria-hidden="true">'
        '<div class="mesh-panel panel-one"></div>'
        '<div class="mesh-panel panel-two"></div>'
        '<div class="mesh-panel panel-three"></div>'
        '<div class="ceiling-lines"></div>'
        '</div>'
        '</div>'
    )

    st.markdown(
        header_html,
        unsafe_allow_html=True,
    )


def render_page_spotlight(page: str):
    page_info = {
        "Yeni Kayıt": (
            "✦",
            "Hızlı üretim girişi",
            "Çalışan, POS ve operasyonları seç; adet, fire ve verim otomatik hesaplansın.",
        ),
        "Operasyon Takibi": (
            "◎",
            "Canlı operasyon görünümü",
            "Her POS'un hangi etapta olduğunu ve kalan miktarını tek bakışta izle.",
        ),
        "Çalışan Yetkinlikleri": (
            "◆",
            "Yetkinlik yönetimi",
            "Çalışanların güçlü olduğu operasyonları kolayca tanımla.",
        ),
        "Verim ve Fire Hedefleri": (
            "◈",
            "Performans hedefleri",
            "Adet/saat, mm²/saat ve fire sınırlarını işlem bazında yönet.",
        ),
        "Üretime Devam Et": (
            "↻",
            "Kaldığın yerden devam",
            "Tamamlanmayan üretimleri ve kalan etapları güvenle sürdür.",
        ),
        "Kayıtlar": (
            "▤",
            "Tüm üretim kayıtları",
            "Kayıt ayrıntılarını incele, indir veya gerektiğinde güvenle sil.",
        ),
        "Grafikler": (
            "▥",
            "Görsel üretim analizi",
            "İlerleme ve operasyon sonuçlarını anlaşılır grafiklerle karşılaştır.",
        ),
        "Yönetici Paneli": (
            "✹",
            "Yönetici kontrol merkezi",
            "Ekip, üretim, mesai, verim ve fire durumunu birlikte değerlendir.",
        ),
        "Gün Sonu Kontrolü": (
            "☑",
            "Gün sonu denetimi",
            "Günün üretim, verim, fire ve mesai sonuçlarını toplu olarak kontrol et.",
        ),
        "Özet": (
            "◉",
            "Genel üretim özeti",
            "OC ve POS bazındaki güncel durumu, tamamlananı ve kalanı görüntüle.",
        ),
    }

    icon, title, copy = page_info.get(
        page,
        ("•", page, "Üretim süreçlerini hızlı ve anlaşılır biçimde yönet."),
    )

    st.markdown(
        f"""
        <div class="page-spotlight">
            <div class="page-spotlight-icon">{icon}</div>
            <div>
                <div class="page-spotlight-title">{title}</div>
                <div class="page-spotlight-copy">{copy}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def render_wizard_stepper(current_step: int):
    steps = [
        (1, "Çalışan & Süre"),
        (2, "OC & POS"),
        (3, "İşlemler"),
        (4, "Adet & Fire"),
        (5, "Kontrol & Onay"),
    ]
    blocks = []
    for number, title in steps:
        state = "done" if number < current_step else "active" if number == current_step else ""
        symbol = "✓" if number < current_step else str(number)
        blocks.append(
            f'<div class="wizard-step {state}">'
            f'<span class="wizard-step-number">{symbol}</span>'
            f'<span class="wizard-step-title">{title}</span>'
            f'</div>'
        )
    st.markdown(
        '<div class="wizard-shell">' + ''.join(blocks) + '</div>',
        unsafe_allow_html=True,
    )


def render_wizard_panel(title: str, copy: str):
    st.markdown(
        f'''
        <div class="wizard-panel">
            <div class="wizard-panel-title">{title}</div>
            <div class="wizard-panel-copy">{copy}</div>
        </div>
        ''',
        unsafe_allow_html=True,
    )


def render_empty_state(title: str, copy: str, icon: str = "◇"):
    st.markdown(
        f'''
        <div class="empty-state">
            <div>
                <div class="empty-icon">{icon}</div>
                <div class="empty-title">{title}</div>
                <div class="empty-copy">{copy}</div>
            </div>
        </div>
        ''',
        unsafe_allow_html=True,
    )


def render_completed_state(title: str, copy: str):
    st.markdown(
        f'''
        <div class="completed-state">
            <strong>✓ {title}</strong><br>
            <span>{copy}</span>
        </div>
        ''',
        unsafe_allow_html=True,
    )


def render_status_pill(label: str, tone: str = "blue"):
    st.markdown(
        f'<span class="status-pill {tone}">{label}</span>',
        unsafe_allow_html=True,
    )


def render_product_footer():
    st.markdown(
        '''
        <div class="product-footer">
            <div><strong>DURLUM FLOW</strong> · Production Intelligence System</div>
            <div>Operasyon · Verim · Fire · Mesai · İzlenebilirlik</div>
        </div>
        ''',
        unsafe_allow_html=True,
    )


def get_manager_notifications() -> list[dict]:
    notifications = []
    today_text = str(date.today())

    try:
        history = get_operation_history()
    except Exception:
        history = pd.DataFrame()

    if not history.empty:
        today = history[history["tarih"].astype(str) == today_text].copy()
        target_map = get_performance_target_map()

        if not today.empty:
            overtime_batches = today[
                today["calisma_tipi"] == "Mesaili"
            ][["batch_id", "operator_ismi", "calisma_saati"]].drop_duplicates("batch_id")
            for _, row in overtime_batches.head(4).iterrows():
                notifications.append({
                    "tone": "purple",
                    "icon": "◷",
                    "title": f"Mesai · {row['operator_ismi']}",
                    "copy": f"Bugün {float(row['calisma_saati']):.1f} saatlik mesaili kayıt oluşturuldu.",
                })

            for _, row in today.iterrows():
                processed = int(row["islem_yapilan"])
                fire = int(row["fire"])
                rate = fire / processed * 100 if processed else 0.0
                limit = float(
                    target_map.get(
                        str(row["operasyon"]),
                        {"fire_limit_pct": float(row.get("fire_siniri_yuzde", 5) or 5)},
                    )["fire_limit_pct"]
                )
                if processed > 0 and rate > limit:
                    notifications.append({
                        "tone": "red",
                        "icon": "!",
                        "title": f"Fire sınırı · {row['POS']} / {row['operasyon']}",
                        "copy": f"Fire %{rate:.1f}; tanımlı sınır %{limit:.1f}.",
                    })

    try:
        batches = get_operation_batch_summary()
    except Exception:
        batches = pd.DataFrame()
    if not batches.empty:
        today_batches = batches[batches["tarih"].astype(str) == today_text]
        for _, row in today_batches[today_batches["gidisat"] == "Yavaş"].head(4).iterrows():
            notifications.append({
                "tone": "orange",
                "icon": "↘",
                "title": f"Yavaş gidişat · {row['operator_ismi']}",
                "copy": f"OC {row['OC']} · {row['POS']} · {float(row['adet_saat']):.2f} adet/saat.",
            })

    try:
        overview = get_operation_overview()
    except Exception:
        overview = pd.DataFrame()
    if not overview.empty:
        near_complete = overview[
            (overview["completion_pct"] >= 85)
            & (overview["completion_pct"] < 100)
        ].sort_values("completion_pct", ascending=False)
        for _, row in near_complete.head(3).iterrows():
            notifications.append({
                "tone": "blue",
                "icon": "◎",
                "title": f"Tamamlanmaya yakın · {row['pos']}",
                "copy": f"OC {row['oc_no']} şu anda %{float(row['completion_pct']):.1f} tamamlandı.",
            })

        completed = overview[overview["completion_pct"] >= 100]
        if not completed.empty:
            notifications.append({
                "tone": "green",
                "icon": "✓",
                "title": "Tamamlanan üretimler",
                "copy": f"Toplam {len(completed)} POS bütün gerekli etapları tamamladı.",
            })

    seen = set()
    unique = []
    for item in notifications:
        key = (item["title"], item["copy"])
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique[:12]


def render_notification_center():
    notifications = get_manager_notifications()
    count = len(notifications)
    label = f"🔔 Bildirim Merkezi · {count}" if count else "🔔 Bildirim Merkezi"
    with st.expander(label, expanded=False):
        if not notifications:
            st.success("Şu anda dikkat gerektiren bir durum bulunmuyor.")
            return
        for item in notifications:
            st.markdown(
                f'''
                <div class="notification-card">
                    <div class="notification-title">
                        <span class="status-pill {item['tone']}">{item['icon']}</span>
                        &nbsp;{item['title']}
                    </div>
                    <div class="notification-copy">{item['copy']}</div>
                </div>
                ''',
                unsafe_allow_html=True,
            )


def render_insight_cards(cards: list[dict]):
    """İçgörü kartlarını HTML sızıntısı olmadan güvenli biçimde gösterir."""
    if not cards:
        return

    for start in range(0, len(cards), 4):
        row_cards = cards[start:start + 4]
        columns = st.columns(len(row_cards))

        for column, card in zip(columns, row_cards):
            icon = str(card.get("icon", "•"))
            title = str(card.get("title", ""))
            copy = str(card.get("copy", ""))

            # HTML'i tek satır halinde veriyoruz. Girintili çok satırlı HTML,
            # bazı Streamlit sürümlerinde Markdown kod bloğuna dönüşebiliyor.
            card_html = (
                '<div class="insight-card">'
                f'<div class="insight-icon">{icon}</div>'
                f'<div class="insight-title">{title}</div>'
                f'<div class="insight-copy">{copy}</div>'
                '</div>'
            )

            with column:
                st.markdown(
                    card_html,
                    unsafe_allow_html=True,
                )



# -----------------------------
# UYGULAMA
# -----------------------------
def main():
    st.set_page_config(
        page_title="DURLUM FLOW · Production Intelligence",
        page_icon="◆",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_durlum_theme()
    init_db()
    init_auth_state()

    is_manager = bool(st.session_state.manager_authenticated)

    render_durlum_header(is_manager)
    st.caption("Sürüm 2.13.1 SIDEBAR FORM OKUNABİLİRLİĞİ · İşçi, OC, POS, işlem, adet, fire ve yetkinlik takibi")
    if is_manager:
        st.success("Yönetici görünümü: Tüm kayıtlar, grafikler, raporlar ve silme işlemleri açık.")
    else:
        st.info("Çalışan görünümü: Yaptığın POS ve işlemleri seçip günlük adetleri girebilirsin.")

    worker_pages = ["Yeni Kayıt", "Üretime Devam Et"]
    manager_pages = [
        "Yeni Kayıt",
        "Operasyon Takibi",
        "Çalışan Yetkinlikleri",
        "Verim ve Fire Hedefleri",
        "Üretime Devam Et",
        "Kayıtlar",
        "Grafikler",
        "Yönetici Paneli",
        "Gün Sonu Kontrolü",
        "Özet",
    ]

    with st.sidebar:
        st.markdown(
            """
            <div class="sidebar-brand">
                <div class="sidebar-mark">d</div>
                <div>
                    <div class="sidebar-brand-title">DURLUM FLOW</div>
                    <div class="sidebar-brand-subtitle">Production Intelligence System</div>
                </div>
            </div>
            <div class="sidebar-version">
                Sürüm 2.13.1 · SIDEBAR FORM OKUNABİLİRLİĞİ
            </div>
            """,
            unsafe_allow_html=True,
        )
        manager_login_panel()

        if is_manager:
            render_notification_center()
        st.divider()

        allowed_pages = manager_pages if is_manager else worker_pages
        page_icons = {
            "Yeni Kayıt": "✦",
            "Operasyon Takibi": "◎",
            "Çalışan Yetkinlikleri": "◆",
            "Verim ve Fire Hedefleri": "◈",
            "Üretime Devam Et": "↻",
            "Kayıtlar": "▤",
            "Grafikler": "▥",
            "Yönetici Paneli": "✹",
            "Gün Sonu Kontrolü": "☑",
            "Özet": "◉",
        }
        page = st.radio(
            "Sayfa",
            allowed_pages,
            index=0,
            format_func=lambda option: f"{page_icons.get(option, '•')}  {option}",
        )

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

    render_page_spotlight(page)

    if page == "Yeni Kayıt":
        if not combinations:
            st.info("Devam etmek için uygulama yöneticisinin ayar Excel'ini tanımlaması gerekir.")
            st.stop()
        yeni_kayit_page(reasons, combinations, is_manager)
    elif page == "Operasyon Takibi" and is_manager:
        operasyon_takibi_page()
    elif page == "Çalışan Yetkinlikleri" and is_manager:
        yetkinlikler_page(combinations)
    elif page == "Verim ve Fire Hedefleri" and is_manager:
        verim_fire_hedefleri_page(combinations)
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

    render_product_footer()


def yeni_kayit_page(reasons, combinations, is_manager: bool):
    st.subheader("Yeni Operasyon Kaydı")
    st.caption(
        "Kayıt beş kısa adımda hazırlanır. Son adımda bütün bilgiler kontrol edilmeden "
        "veritabanına yazılmaz."
    )

    success_message = st.session_state.pop("operation_entry_success", None)
    if success_message:
        st.success(success_message)

    all_items = get_production_output_summary()
    if all_items.empty:
        render_empty_state(
            "Henüz üretim çıktısı yüklenmedi",
            "Yönetici sol menüden PDF veya Excel üretim çıktısını yüklediğinde çalışanlar "
            "OC ve POS seçerek operasyona başlayabilir.",
            "⇧",
        )
        return

    version = int(st.session_state.get("operation_form_version", 0))
    step_key = f"operation_wizard_step_{version}"
    if step_key not in st.session_state:
        st.session_state[step_key] = 1
    step = int(st.session_state[step_key])
    render_wizard_stepper(step)

    worker_key = f"op_worker_{version}"
    date_key = f"op_date_{version}"
    type_key = f"op_work_type_{version}"
    reason_key = f"op_reason_{version}"
    oc_key = f"op_oc_{version}"

    # Streamlit, bir adımda gösterilmeyen widget anahtarlarını temizleyebilir.
    # Bu nedenle sihirbaz bilgilerini widget'lardan bağımsız kalıcı bir sözlükte tutuyoruz.
    wizard_data_key = f"operation_wizard_data_{version}"
    if wizard_data_key not in st.session_state:
        st.session_state[wizard_data_key] = {}
    wizard_data = dict(st.session_state.get(wizard_data_key, {}))

    def save_wizard_data(**values):
        current = dict(st.session_state.get(wizard_data_key, {}))
        current.update(values)
        st.session_state[wizard_data_key] = current
        wizard_data.update(values)

    # Önceki adımdan geri dönüldüğünde widget'ları kalıcı verilerle tekrar doldur.
    if worker_key not in st.session_state and wizard_data.get("operator_name") is not None:
        st.session_state[worker_key] = wizard_data.get("operator_name", "")
    if date_key not in st.session_state and wizard_data.get("production_date") is not None:
        st.session_state[date_key] = wizard_data.get("production_date", date.today())
    if type_key not in st.session_state and wizard_data.get("work_type") is not None:
        st.session_state[type_key] = wizard_data.get("work_type", "Tam zamanlı")
    if reason_key not in st.session_state and wizard_data.get("reason") is not None:
        st.session_state[reason_key] = wizard_data.get("reason", "")

    oc_table = (
        all_items[["oc_no", "project_name"]]
        .drop_duplicates()
        .sort_values(["oc_no", "project_name"])
        .reset_index(drop=True)
    )
    oc_labels = oc_table.apply(
        lambda row: f"OC {row['oc_no']} — {row['project_name']}", axis=1
    ).tolist()
    oc_map = dict(zip(oc_labels, oc_table["oc_no"].astype(str)))

    def go_to(target_step: int):
        st.session_state[step_key] = int(target_step)
        st.rerun()

    def current_work_hours() -> float:
        current_type = str(
            st.session_state.get(
                type_key,
                wizard_data.get("work_type", "Tam zamanlı"),
            )
        )
        hours_key = f"op_hours_{current_type}_{version}"
        defaults = {"Tam zamanlı": 9.0, "Yarı zamanlı": 6.0, "Mesaili": 10.0}
        if hours_key in st.session_state:
            return float(st.session_state[hours_key])
        return float(
            wizard_data.get(
                "work_hours",
                defaults.get(current_type, 9.0),
            )
        )

    def resolve_selected_context():
        selected_label = st.session_state.get(
            oc_key,
            wizard_data.get(
                "selected_oc_label",
                oc_labels[0] if oc_labels else "",
            ),
        )
        selected_oc = oc_map.get(
            selected_label,
            wizard_data.get(
                "selected_oc",
                oc_table["oc_no"].astype(str).iloc[0],
            ),
        )
        oc_items = get_production_output_summary(selected_oc)
        positions_key = f"op_positions_{selected_oc}_{version}"
        positions = list(
            st.session_state.get(
                positions_key,
                wizard_data.get("selected_positions", []),
            )
        )
        selected_items = oc_items[
            oc_items["pos"].astype(str).isin(positions)
        ].copy()
        if not selected_items.empty:
            selected_items["_sort_pos"] = (
                selected_items["pos"].astype(str).str.extract(r"(\d+)")[0]
                .fillna(0).astype(int)
            )
        return selected_oc, oc_items, positions, selected_items

    combo_map = {combo["ad"]: list(combo["operasyonlar"]) for combo in combinations}
    combo_names = list(combo_map)

    def resolve_plans(selected_oc: str, selected_items: pd.DataFrame):
        existing_plans = {}
        missing_plan_ids = []
        for _, item in selected_items.iterrows():
            item_id = int(item["item_id"])
            plan_df = get_item_operation_plan(item_id)
            if plan_df.empty:
                missing_plan_ids.append(item_id)
            else:
                existing_plans[item_id] = {
                    "combination_name": str(plan_df["combination_name"].iloc[0]),
                    "operations": plan_df["operation_name"].astype(str).tolist(),
                }
        change_key = f"change_plan_{selected_oc}_{version}"
        change_plan = (
            bool(
                st.session_state.get(
                    change_key,
                    wizard_data.get("change_plan", False),
                )
            )
            if is_manager
            else False
        )
        combo_key = f"op_combo_{selected_oc}_{version}"
        selected_combo = st.session_state.get(
            combo_key,
            wizard_data.get(
                "selected_combo",
                combo_names[0] if combo_names else None,
            ),
        )
        plans = {}
        for _, item in selected_items.iterrows():
            item_id = int(item["item_id"])
            if change_plan or item_id in missing_plan_ids:
                plans[item_id] = {
                    "combination_name": selected_combo,
                    "operations": combo_map.get(selected_combo, []),
                }
            else:
                plans[item_id] = existing_plans[item_id]
        return plans, existing_plans, missing_plan_ids, change_plan, selected_combo

    def resolve_operation_choices(
        selected_oc: str,
        selected_items: pd.DataFrame,
        plans: dict,
        operator_name: str,
    ):
        competencies = get_worker_competencies(operator_name)
        show_other_key = f"show_other_ops_{operator_name}_{version}"
        show_other = bool(
            st.session_state.get(
                show_other_key,
                wizard_data.get("show_other_operations", False),
            )
        )

        def allowed(plan_operations):
            if not competencies or show_other:
                return list(plan_operations)
            return [operation for operation in plan_operations if operation in competencies]

        same_key = f"same_operations_{selected_oc}_{version}"
        same_operations = bool(
            st.session_state.get(
                same_key,
                wizard_data.get("same_operations", True),
            )
        )
        saved_selected_map = {
            int(item_id): list(operations)
            for item_id, operations in wizard_data.get(
                "selected_operation_map",
                {},
            ).items()
        }
        selected_map = {}
        if same_operations:
            common_key = f"common_ops_{selected_oc}_{version}"
            common = list(
                st.session_state.get(
                    common_key,
                    wizard_data.get("common_operations", []),
                )
            )
            for _, item in selected_items.iterrows():
                item_id = int(item["item_id"])
                available = allowed(plans[item_id]["operations"])
                if common:
                    selected_map[item_id] = [
                        operation for operation in common
                        if operation in available
                    ]
                else:
                    selected_map[item_id] = [
                        operation for operation in saved_selected_map.get(item_id, [])
                        if operation in available
                    ]
        else:
            for _, item in selected_items.iterrows():
                item_id = int(item["item_id"])
                ops_key = f"ops_{item_id}_{version}"
                selected_map[item_id] = list(
                    st.session_state.get(
                        ops_key,
                        saved_selected_map.get(item_id, []),
                    )
                )
        return selected_map, competencies, show_other, allowed

    def build_entries(selected_items: pd.DataFrame, plans: dict, selected_map: dict):
        if selected_items.empty:
            return []
        progress = get_operation_progress(selected_items["item_id"].astype(int).tolist())
        lookup = {}
        if not progress.empty:
            for _, row in progress.iterrows():
                lookup[(int(row["item_id"]), str(row["operation_name"]))] = row
        entries = []
        for _, item in selected_items.iterrows():
            item_id = int(item["item_id"])
            requested = int(item["requested_qty"])
            for operation in selected_map.get(item_id, []):
                progress_row = lookup.get((item_id, operation))
                completed = int(progress_row["completed_qty"]) if progress_row is not None else 0
                remaining = max(requested - completed, 0)
                processed_key = f"processed_{item_id}_{_normalize_header(operation)}_{version}"
                fire_key = f"opfire_{item_id}_{_normalize_header(operation)}_{version}"

                saved_entry_map = {
                    (
                        int(saved_entry["item_id"]),
                        str(saved_entry["operation_name"]),
                    ): saved_entry
                    for saved_entry in wizard_data.get("entries", [])
                }
                saved_entry = saved_entry_map.get((item_id, operation), {})

                processed = int(
                    st.session_state.get(
                        processed_key,
                        saved_entry.get("processed_qty", 0),
                    )
                )
                fire = int(
                    st.session_state.get(
                        fire_key,
                        saved_entry.get("fire_qty", 0),
                    )
                )
                if processed > 0:
                    entries.append({
                        "item_id": item_id,
                        "pos": str(item["pos"]),
                        "operation_name": operation,
                        "processed_qty": processed,
                        "fire_qty": fire,
                        "good_qty": processed - fire,
                        "remaining_before": remaining,
                        "unit_area_mm2": float(item["unit_area_mm2"]),
                    })
        return entries

    # --------------------------------------------------------------
    # ADIM 1
    # --------------------------------------------------------------
    if step == 1:
        render_wizard_panel(
            "1 · Çalışan ve çalışma süresi",
            "Kayıt sahibini, tarihi ve o günün toplam çalışma süresini seç.",
        )
        c1, c2, c3 = st.columns(3)
        with c1:
            production_date = st.date_input(
                "Tarih",
                value=wizard_data.get("production_date", date.today()),
                key=date_key,
            )
        with c2:
            operator_name = st.selectbox(
                "Operatör / İşçi",
                WORKER_NAMES,
                format_func=lambda value: "İsim seçin" if value == "" else value,
                key=worker_key,
            )
        with c3:
            work_type = st.selectbox(
                "Çalışma tipi",
                ["Tam zamanlı", "Yarı zamanlı", "Mesaili"],
                key=type_key,
            )

        defaults = {"Tam zamanlı": 9.0, "Yarı zamanlı": 6.0, "Mesaili": 10.0}
        hours_widget_key = f"op_hours_{work_type}_{version}"
        if hours_widget_key not in st.session_state:
            st.session_state[hours_widget_key] = float(
                wizard_data.get("work_hours", defaults[work_type])
            )
        h1, h2 = st.columns(2)
        with h1:
            work_hours = st.number_input(
                "O gün toplam kaç saat çalıştı?",
                min_value=0.0,
                max_value=24.0,
                step=0.5,
                key=hours_widget_key,
            )
        with h2:
            reason = st.selectbox(
                "Yarı zamanlı çalışma nedeni",
                [""] + reasons,
                disabled=(work_type != "Yarı zamanlı"),
                key=reason_key,
            )

        if work_type == "Mesaili":
            render_status_pill("◷ Mesaili kayıt · Yönetici bildirimi oluşur", "purple")

        if st.button("OC ve POS Seçimine Geç →", type="primary", use_container_width=True):
            errors = []
            if not operator_name:
                errors.append("Operatör / işçi seçmelisin.")
            if float(work_hours) <= 0:
                errors.append("Çalışma süresi 0'dan büyük olmalıdır.")
            if work_type == "Yarı zamanlı" and not reason:
                errors.append("Yarı zamanlı çalışma için neden seçmelisin.")
            if work_type == "Mesaili" and float(work_hours) <= 9.0:
                errors.append("Mesaili çalışma süresi 9 saatten fazla olmalıdır.")
            if errors:
                for error in errors:
                    st.error(error)
            else:
                save_wizard_data(
                    production_date=production_date,
                    operator_name=operator_name,
                    work_type=work_type,
                    work_hours=float(work_hours),
                    reason=reason,
                )
                go_to(2)
        return

    # --------------------------------------------------------------
    # ADIM 2
    # --------------------------------------------------------------
    if step == 2:
        render_wizard_panel(
            "2 · OC ve POS seçimi",
            "Aynı gün birden fazla POS üzerinde çalışıldıysa hepsini birlikte seçebilirsin.",
        )
        if oc_key not in st.session_state:
            st.session_state[oc_key] = wizard_data.get(
                "selected_oc_label",
                oc_labels[0],
            )
        selected_oc_label = st.selectbox("OC seç", oc_labels, key=oc_key)
        selected_oc = oc_map[selected_oc_label]
        oc_items = get_production_output_summary(selected_oc)
        pos_options = oc_items["pos"].astype(str).tolist()
        positions_key = f"op_positions_{selected_oc}_{version}"
        if positions_key not in st.session_state:
            st.session_state[positions_key] = [
                position
                for position in wizard_data.get("selected_positions", [])
                if position in pos_options
            ]
        selected_positions = st.multiselect(
            "Bugün çalışılan POS veya POS'ları seç",
            pos_options,
            key=positions_key,
            help="POS10, POS30 ve diğer çalışılan POS'ları birlikte seçebilirsin.",
        )

        selected_items = oc_items[
            oc_items["pos"].astype(str).isin(selected_positions)
        ].copy()
        plans, existing_plans, missing_ids, change_plan, selected_combo = resolve_plans(
            selected_oc, selected_items
        )

        if is_manager and existing_plans:
            change_plan = st.checkbox(
                "Seçilen POS'ların kombinasyonunu yeniden ata",
                value=change_plan,
                key=f"change_plan_{selected_oc}_{version}",
            )

        if selected_positions and (missing_ids or change_plan):
            selected_combo = st.selectbox(
                "Kombinasyon / gerekli etaplar",
                combo_names,
                key=f"op_combo_{selected_oc}_{version}",
            )
            st.caption("Etap akışı: " + " → ".join(combo_map[selected_combo]))

        if selected_positions:
            chosen = oc_items[oc_items["pos"].astype(str).isin(selected_positions)]
            total_requested = int(chosen["requested_qty"].sum())
            total_remaining = int(chosen["remaining_qty"].sum())
            p1, p2, p3 = st.columns(3)
            p1.metric("Seçilen POS", len(selected_positions))
            p2.metric("Toplam istenen", total_requested)
            p3.metric("Tam bitmemiş adet", total_remaining)

        b1, b2 = st.columns(2)
        with b1:
            if st.button("← Çalışan Bilgilerine Dön", use_container_width=True):
                go_to(1)
        with b2:
            if st.button("İşlem Seçimine Geç →", type="primary", use_container_width=True):
                if not selected_positions:
                    st.error("En az bir POS seçmelisin.")
                elif (missing_ids or change_plan) and not selected_combo:
                    st.error("Kombinasyon seçmelisin.")
                else:
                    save_wizard_data(
                        selected_oc_label=selected_oc_label,
                        selected_oc=selected_oc,
                        selected_positions=list(selected_positions),
                        selected_combo=selected_combo,
                        change_plan=bool(change_plan),
                    )
                    go_to(3)
        return

    selected_oc, oc_items, selected_positions, selected_items = resolve_selected_context()
    if not selected_positions or selected_items.empty:
        st.warning("OC/POS seçimi bulunamadı. İkinci adıma dönülüyor.")
        go_to(2)
    plans, existing_plans, missing_ids, change_plan, selected_combo = resolve_plans(
        selected_oc, selected_items
    )
    operator_name = str(
        st.session_state.get(
            worker_key,
            wizard_data.get("operator_name", ""),
        )
    )

    # --------------------------------------------------------------
    # ADIM 3
    # --------------------------------------------------------------
    if step == 3:
        render_wizard_panel(
            "3 · Yapılan işlemler",
            "Çalışanın o gün gerçekten yaptığı etapları seç. Tamamlanmış etapları yeniden seçme.",
        )
        all_plan_operations = []
        for plan in plans.values():
            for operation in plan["operations"]:
                if operation not in all_plan_operations:
                    all_plan_operations.append(operation)

        competencies = get_worker_competencies(operator_name)
        show_other = True
        if competencies:
            st.info("Tanımlı yetkinlikler: " + ", ".join(competencies))
            show_other = st.checkbox(
                "Yetkinliklerim dışındaki işlemleri de göster",
                value=bool(st.session_state.get(f"show_other_ops_{operator_name}_{version}", False)),
                key=f"show_other_ops_{operator_name}_{version}",
            )
        else:
            st.caption("Bu çalışan için yetkinlik sınırı tanımlanmamış; bütün plan işlemleri gösteriliyor.")

        def allowed_for_plan(plan_operations):
            if not competencies or show_other:
                return list(plan_operations)
            return [operation for operation in plan_operations if operation in competencies]

        same_operations_key = f"same_operations_{selected_oc}_{version}"
        if same_operations_key not in st.session_state:
            st.session_state[same_operations_key] = bool(
                wizard_data.get("same_operations", True)
            )
        same_operations = st.checkbox(
            "Seçtiğim POS'larda aynı işlemleri yaptım",
            key=same_operations_key,
        )

        if same_operations:
            operation_sets = [set(allowed_for_plan(plan["operations"])) for plan in plans.values()]
            common_options = [
                operation for operation in all_plan_operations
                if all(operation in operation_set for operation_set in operation_sets)
            ]
            if not common_options:
                st.warning("Seçilen POS'ların ortak işlemi yok. Aynı işlemler seçimini kapat.")
            else:
                common_operations_key = f"common_ops_{selected_oc}_{version}"
                if common_operations_key not in st.session_state:
                    st.session_state[common_operations_key] = [
                        operation
                        for operation in wizard_data.get("common_operations", [])
                        if operation in common_options
                    ]
                st.multiselect(
                    "Bugün yapılan işlem veya işlemler",
                    common_options,
                    key=common_operations_key,
                )
        else:
            for _, item in selected_items.sort_values("_sort_pos").iterrows():
                item_id = int(item["item_id"])
                per_pos_key = f"ops_{item_id}_{version}"
                available_operations = allowed_for_plan(
                    plans[item_id]["operations"]
                )
                if per_pos_key not in st.session_state:
                    saved_map = wizard_data.get(
                        "selected_operation_map",
                        {},
                    )
                    saved_operations = (
                        saved_map.get(item_id)
                        or saved_map.get(str(item_id))
                        or []
                    )
                    st.session_state[per_pos_key] = [
                        operation
                        for operation in saved_operations
                        if operation in available_operations
                    ]
                st.multiselect(
                    f"{item['pos']} için bugün yapılan işlemler",
                    available_operations,
                    key=per_pos_key,
                )

        selected_map, _, _, _ = resolve_operation_choices(
            selected_oc, selected_items, plans, operator_name
        )
        selected_count = sum(len(values) for values in selected_map.values())
        if selected_count:
            render_status_pill(f"✓ {selected_count} POS/işlem seçimi hazır", "green")

        b1, b2 = st.columns(2)
        with b1:
            if st.button("← OC ve POS'a Dön", use_container_width=True):
                go_to(2)
        with b2:
            if st.button("Adet ve Fire Girişine Geç →", type="primary", use_container_width=True):
                if selected_count <= 0:
                    st.error("En az bir işlem seçmelisin.")
                else:
                    common_operations = []
                    if same_operations:
                        common_operations = list(
                            st.session_state.get(
                                f"common_ops_{selected_oc}_{version}",
                                [],
                            )
                        )
                    save_wizard_data(
                        same_operations=bool(same_operations),
                        show_other_operations=bool(show_other),
                        common_operations=common_operations,
                        selected_operation_map={
                            str(item_id): list(operations)
                            for item_id, operations in selected_map.items()
                        },
                    )
                    go_to(4)
        return

    selected_map, competencies, show_other, allowed_for_plan = resolve_operation_choices(
        selected_oc, selected_items, plans, operator_name
    )

    # --------------------------------------------------------------
    # ADIM 4
    # --------------------------------------------------------------
    if step == 4:
        render_wizard_panel(
            "4 · Adet ve fire",
            "Her seçilen işlem için toplam işlenen adedi ve varsa fireyi gir. Sağlam ilerleme otomatik hesaplanır.",
        )
        progress = get_operation_progress(selected_items["item_id"].astype(int).tolist())
        progress_lookup = {}
        if not progress.empty:
            for _, row in progress.iterrows():
                progress_lookup[(int(row["item_id"]), str(row["operation_name"]))] = row

        for _, item in selected_items.sort_values("_sort_pos").iterrows():
            item_id = int(item["item_id"])
            pos_name = str(item["pos"])
            requested_qty = int(item["requested_qty"])
            selected_ops = selected_map.get(item_id, [])
            if not selected_ops:
                continue

            with st.expander(
                f"{pos_name} · İstenen {requested_qty} adet · {plans[item_id]['combination_name']}",
                expanded=True,
            ):
                render_pos_area_overview(
                    boy_mm=float(item["boy_mm"]),
                    en_mm=float(item["en_mm"]),
                    unit_area_mm2=float(item["unit_area_mm2"]),
                    requested_qty=requested_qty,
                    completed_qty=int(item["produced_qty"]),
                )

                for operation in selected_ops:
                    progress_row = progress_lookup.get((item_id, operation))
                    completed_qty = int(progress_row["completed_qty"]) if progress_row is not None else 0
                    remaining_qty = max(requested_qty - completed_qty, 0)
                    processed_key = f"processed_{item_id}_{_normalize_header(operation)}_{version}"
                    fire_key = f"opfire_{item_id}_{_normalize_header(operation)}_{version}"

                    saved_entry_lookup = {
                        (
                            int(saved_entry["item_id"]),
                            str(saved_entry["operation_name"]),
                        ): saved_entry
                        for saved_entry in wizard_data.get("entries", [])
                    }
                    saved_entry = saved_entry_lookup.get(
                        (item_id, operation),
                        {},
                    )
                    if processed_key not in st.session_state:
                        st.session_state[processed_key] = int(
                            saved_entry.get("processed_qty", 0)
                        )
                    if fire_key not in st.session_state:
                        st.session_state[fire_key] = int(
                            saved_entry.get("fire_qty", 0)
                        )

                    row1, row2, row3, row4 = st.columns([2.2, 1.1, 1.0, 1.2])
                    with row1:
                        st.markdown(f"**{operation}**")
                        tone = "green" if remaining_qty == 0 else "blue" if completed_qty > 0 else "gray"
                        render_status_pill(
                            f"{completed_qty}/{requested_qty} tamamlandı · {remaining_qty} kaldı",
                            tone,
                        )
                    with row2:
                        st.number_input(
                            "İşlem yapılan",
                            min_value=0,
                            step=1,
                            key=processed_key,
                        )
                    with row3:
                        st.number_input(
                            "Fire",
                            min_value=0,
                            step=1,
                            key=fire_key,
                        )
                    with row4:
                        st.write("")
                        st.button(
                            "Kalanı yaz",
                            disabled=(remaining_qty == 0),
                            use_container_width=True,
                            key=f"fill_{item_id}_{_normalize_header(operation)}_{version}",
                            on_click=set_session_value,
                            args=(processed_key, remaining_qty),
                        )
                    st.divider()

        entries = build_entries(selected_items, plans, selected_map)
        performance = calculate_entries_performance(entries, current_work_hours())
        m1, m2, m3 = st.columns(3)
        m1.metric("İşlem yapılan", f"{performance['total_processed']} adet")
        m2.metric("Sağlam ilerleyen", f"{performance['total_good']} adet")
        m3.metric("Fire", f"{performance['total_fire']} adet")

        render_area_summary(performance, entries)

        v1, v2, v3 = st.columns(3)
        v1.metric("Adet / saat", f"{performance['qty_per_hour']:.2f}")
        v2.metric("Fire oranı", f"%{performance['fire_rate_pct']:.2f}")
        v3.metric(
            "Gidişat",
            performance["status"],
            delta=(
                f"%{performance['score_pct']:.1f}"
                if performance["score_pct"] is not None
                else performance["reference_source"]
            ),
        )

        if performance["status"] == "Yavaş":
            st.warning("Gidişat referansa göre yavaş görünüyor.")
        elif performance["status"] == "Hızlı":
            st.success("Gidişat referansa göre hızlı görünüyor.")
        elif performance["status"] == "Normal":
            st.info("Gidişat hedeflenen normal aralıkta.")

        fire_stage_key = f"operation_fire_stage_{version}"
        fire_note_key = f"operation_fire_note_{version}"
        note_key = f"operation_note_{version}"

        fire_stage_options = sorted(
            {
                str(entry["operation_name"])
                for entry in entries
                if int(entry.get("fire_qty", 0)) > 0
            }
        )
        if fire_stage_key not in st.session_state:
            saved_fire_stage = str(wizard_data.get("fire_stage", ""))
            st.session_state[fire_stage_key] = (
                saved_fire_stage
                if saved_fire_stage in fire_stage_options
                else ""
            )
        if fire_note_key not in st.session_state:
            st.session_state[fire_note_key] = wizard_data.get(
                "fire_note",
                "",
            )
        if note_key not in st.session_state:
            st.session_state[note_key] = wizard_data.get(
                "note",
                "",
            )
        fire_col1, fire_col2 = st.columns([1, 2])
        with fire_col1:
            st.selectbox(
                "Fire hangi etapta oluştu?",
                [""] + fire_stage_options,
                disabled=(int(performance["total_fire"]) == 0),
                key=fire_stage_key,
                format_func=lambda value: (
                    "Etap seçin" if value == "" else value
                ),
            )
        with fire_col2:
            st.text_input(
                "Fire açıklaması",
                disabled=(int(performance["total_fire"]) == 0),
                key=fire_note_key,
                placeholder="Örnek: Parçada yüzey hatası oluştu.",
            )
        st.text_area("Genel not", key=note_key)

        for breach in performance["fire_breaches"]:
            st.error(
                f"{breach['pos']} / {breach['operation']}: Fire %{breach['fire_rate_pct']:.1f}; "
                f"sınır %{breach['fire_limit_pct']:.1f}."
            )

        b1, b2 = st.columns(2)
        with b1:
            if st.button("← İşlemlere Dön", use_container_width=True):
                go_to(3)
        with b2:
            if st.button("Kayıt Özetini Gör →", type="primary", use_container_width=True):
                errors = []
                if not entries:
                    errors.append("En az bir işlem için yapılan adet girmelisin.")
                for entry in entries:
                    if entry["fire_qty"] > entry["processed_qty"]:
                        errors.append(
                            f"{entry['pos']} / {entry['operation_name']}: Fire işlenen adetten büyük olamaz."
                        )
                    if entry["good_qty"] > entry["remaining_before"]:
                        errors.append(
                            f"{entry['pos']} / {entry['operation_name']}: Sağlam ilerleyen adet kalan miktarı aşamaz."
                        )
                if int(performance["total_fire"]) > 0:
                    if not str(
                        st.session_state.get(fire_stage_key, "")
                    ).strip():
                        errors.append(
                            "Fire varsa hangi etapta oluştuğunu seçmelisin."
                        )
                    if not str(
                        st.session_state.get(fire_note_key, "")
                    ).strip():
                        errors.append(
                            "Fire varsa kısa bir açıklama yazmalısın."
                        )
                if errors:
                    for error in errors:
                        st.error(error)
                else:
                    save_wizard_data(
                        entries=[dict(entry) for entry in entries],
                        fire_stage=str(
                            st.session_state.get(fire_stage_key, "")
                        ),
                        fire_note=str(
                            st.session_state.get(fire_note_key, "")
                        ),
                        note=str(
                            st.session_state.get(note_key, "")
                        ),
                    )
                    go_to(5)
        return

    # --------------------------------------------------------------
    # ADIM 5
    # --------------------------------------------------------------
    render_wizard_panel(
        "5 · Kontrol ve onay",
        "Aşağıdaki özet veritabanına yazılacak son kayıttır. Onaylamadan önce adetleri ve işlemleri kontrol et.",
    )
    entries = [
        dict(entry)
        for entry in wizard_data.get("entries", [])
    ]
    if not entries:
        entries = build_entries(selected_items, plans, selected_map)

    work_type = str(
        wizard_data.get(
            "work_type",
            st.session_state.get(type_key, "Tam zamanlı"),
        )
    )
    production_date = wizard_data.get(
        "production_date",
        st.session_state.get(date_key, date.today()),
    )
    reason = str(
        wizard_data.get(
            "reason",
            st.session_state.get(reason_key, ""),
        )
    )
    work_hours = float(
        wizard_data.get(
            "work_hours",
            current_work_hours(),
        )
    )
    operator_name = str(
        wizard_data.get(
            "operator_name",
            operator_name,
        )
    )
    selected_oc = str(
        wizard_data.get(
            "selected_oc",
            selected_oc,
        )
    )
    selected_positions = list(
        wizard_data.get(
            "selected_positions",
            selected_positions,
        )
    )

    performance = calculate_entries_performance(entries, work_hours)
    fire_stage = str(
        wizard_data.get(
            "fire_stage",
            st.session_state.get(
                f"operation_fire_stage_{version}",
                "",
            ),
        )
    ).strip()
    fire_note = str(
        wizard_data.get(
            "fire_note",
            st.session_state.get(
                f"operation_fire_note_{version}",
                "",
            ),
        )
    ).strip()
    note = str(
        wizard_data.get(
            "note",
            st.session_state.get(
                f"operation_note_{version}",
                "",
            ),
        )
    ).strip()

    status_tone = {
        "Yavaş": "orange",
        "Normal": "blue",
        "Hızlı": "green",
        "Referans yok": "gray",
    }.get(performance["status"], "blue")

    st.markdown(
        f'''
        <div class="review-card">
            <div class="review-title">Kayıt Özeti</div>
            <div class="review-grid">
                <div class="review-item"><div class="review-label">Çalışan</div><div class="review-value">{operator_name}</div></div>
                <div class="review-item"><div class="review-label">Tarih</div><div class="review-value">{production_date}</div></div>
                <div class="review-item"><div class="review-label">Çalışma</div><div class="review-value">{work_type} · {work_hours:.1f} saat</div></div>
                <div class="review-item"><div class="review-label">OC / POS</div><div class="review-value">{selected_oc} · {len(selected_positions)} POS</div></div>
                <div class="review-item"><div class="review-label">İşlem yapılan</div><div class="review-value">{performance['total_processed']} adet</div></div>
                <div class="review-item"><div class="review-label">Sağlam ilerleyen</div><div class="review-value">{performance['total_good']} adet</div></div>
                <div class="review-item"><div class="review-label">Fire</div><div class="review-value">{performance['total_fire']} adet · %{performance['fire_rate_pct']:.2f}</div></div>
                <div class="review-item"><div class="review-label">Verim</div><div class="review-value">{performance['qty_per_hour']:.2f} adet/saat</div></div>
            </div>
        </div>
        ''',
        unsafe_allow_html=True,
    )
    render_status_pill(
        f"Gidişat: {performance['status']} · {performance['reference_source']}",
        status_tone,
    )

    summary_rows = []
    for entry in entries:
        summary_rows.append({
            "POS": entry["pos"],
            "İşlem": entry["operation_name"],
            "İşlenen": entry["processed_qty"],
            "Fire": entry["fire_qty"],
            "Sağlam": entry["good_qty"],
            "Önceki Kalan": entry["remaining_before"],
            "Kayıt Sonrası Kalan": max(entry["remaining_before"] - entry["good_qty"], 0),
            "Sağlam Alan (m²)": round(
                entry["good_qty"] * entry["unit_area_mm2"] / 1_000_000,
                3,
            ),
        })
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    render_area_summary(performance, entries)

    if fire_note:
        st.warning(
            f"Fire etabı: {fire_stage or '-'} · Açıklama: {fire_note}"
        )
    if note:
        st.info("Genel not: " + note)

    confirm_key = f"operation_final_confirm_{version}"
    confirmed = st.checkbox(
        "Bilgileri kontrol ettim; kayıt özeti doğrudur.",
        key=confirm_key,
    )

    b1, b2 = st.columns([1, 1.35])
    with b1:
        if st.button("← Düzenlemeye Dön", use_container_width=True):
            go_to(4)
    with b2:
        save_label = "Kaydı Onayla ve Tamamla ✓"
        if st.button(
            save_label,
            type="primary",
            use_container_width=True,
            disabled=not confirmed,
        ):
            final_errors = []
            if not operator_name:
                final_errors.append("Operatör / işçi seçimi bulunamadı.")
            if not entries:
                final_errors.append("Kaydedilecek işlem adedi bulunamadı.")
            if work_type == "Yarı zamanlı" and not reason:
                final_errors.append("Yarı zamanlı çalışma için neden seçmelisin.")
            if work_type == "Mesaili" and float(work_hours) <= 9.0:
                final_errors.append("Mesaili çalışma süresi 9 saatten fazla olmalıdır.")
            for entry in entries:
                if entry["fire_qty"] > entry["processed_qty"]:
                    final_errors.append(
                        f"{entry['pos']} / {entry['operation_name']}: Fire işlenen adetten büyük olamaz."
                    )
                if entry["good_qty"] > entry["remaining_before"]:
                    final_errors.append(
                        f"{entry['pos']} / {entry['operation_name']}: Sağlam ilerleyen adet kalan miktarı aşamaz."
                    )
            if performance["total_fire"] > 0:
                if not fire_stage:
                    final_errors.append(
                        "Fire varsa hangi etapta oluştuğunu seçmelisin."
                    )
                if not fire_note:
                    final_errors.append(
                        "Fire varsa kısa bir açıklama yazmalısın."
                    )

            if final_errors:
                for error in final_errors:
                    st.error(error)
                st.stop()

            for entry in entries:
                entry["fire_note"] = (
                    f"Fire etabı: {fire_stage} | Açıklama: {fire_note}"
                    if entry["fire_qty"] > 0
                    else ""
                )

            batch_id = save_operation_batch(
                {
                    "tarih": str(production_date),
                    "operator_ismi": operator_name.strip(),
                    "calisma_tipi": work_type,
                    "calisma_saati": float(work_hours),
                    "neden": reason,
                    "notlar": note,
                },
                entries,
                plans,
            )
            st.session_state.operation_form_version = version + 1
            st.session_state.operation_entry_success = (
                f"Kayıt tamamlandı. {len(entries)} POS/işlem satırı kaydedildi. Kayıt no: {batch_id}."
            )
            st.rerun()



def yetkinlikler_page(combinations):
    st.subheader("Çalışan Yetkinlikleri")
    st.caption(
        "Çalışanın sık yaptığı işlemleri tanımla. Bu bir engel değildir; çalışan gerektiğinde diğer işlemleri de gösterebilir."
    )
    all_operations = all_operations_from_combinations(combinations)
    worker = st.selectbox(
        "Çalışan",
        [name for name in WORKER_NAMES if name],
        key="competency_worker",
    )
    current = get_worker_competencies(worker)
    selected = st.multiselect(
        "Yetkin olduğu işlemler",
        all_operations,
        default=[operation for operation in current if operation in all_operations],
        key=f"competencies_{worker}",
    )
    if st.button("Yetkinlikleri Kaydet", type="primary"):
        save_worker_competencies(worker, selected)
        st.success(f"{worker} için yetkinlikler kaydedildi.")

    table = get_competency_table()
    if table.empty:
        st.info("Henüz yetkinlik tanımı yok. Tanımsız çalışanlarda bütün işlemler gösterilir.")
    else:
        summary = (
            table.groupby("worker_name")["operation_name"]
            .apply(lambda values: ", ".join(values))
            .reset_index()
            .rename(columns={"worker_name": "Çalışan", "operation_name": "Yetkinlikler"})
        )
        st.dataframe(summary, use_container_width=True, hide_index=True)


def operasyon_takibi_page():
    st.subheader("Operasyon Takibi")
    st.caption(
        "Her POS'un lazer, kaynak, boya ve diğer etaplarda kaç adede ulaştığını gösterir. "
        "Tam biten adet, gerekli etaplar içindeki en düşük ilerlemedir."
    )

    overview = get_operation_overview()
    if overview.empty:
        render_empty_state("Takip edilecek üretim çıktısı yok", "Yönetici bir üretim çıktısı yüklediğinde OC, POS ve etap ilerlemeleri burada görünür.", "◎")
        return

    f1, f2 = st.columns(2)
    oc_options = ["Tümü"] + sorted(overview["oc_no"].astype(str).unique().tolist())
    with f1:
        selected_oc = st.selectbox("OC", oc_options, key="operation_tracking_oc")
    filtered = overview.copy()
    if selected_oc != "Tümü":
        filtered = filtered[filtered["oc_no"].astype(str) == selected_oc]
    pos_options = ["Tümü"] + filtered["pos"].astype(str).tolist()
    with f2:
        selected_pos = st.selectbox("POS", pos_options, key="operation_tracking_pos")
    if selected_pos != "Tümü":
        filtered = filtered[filtered["pos"].astype(str) == selected_pos]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("İstenen", int(filtered["requested_qty"].sum()))
    k2.metric("Tam biten", int(filtered["produced_qty"].sum()))
    k3.metric("Üretimde olan", int(filtered["in_production_qty"].sum()))
    k4.metric("Operasyon firesi", int(filtered["operation_fire_qty"].sum()))

    summary = filtered[[
        "oc_no", "pos", "combination_name", "requested_qty", "produced_qty",
        "in_production_qty", "remaining_qty", "operation_fire_qty", "completion_pct",
    ]].rename(columns={
        "oc_no": "OC",
        "pos": "POS",
        "combination_name": "Kombinasyon",
        "requested_qty": "İstenen",
        "produced_qty": "Tam Biten",
        "in_production_qty": "Üretimde",
        "remaining_qty": "Kalan",
        "operation_fire_qty": "Fire",
        "completion_pct": "Tamamlanma %",
    })
    st.dataframe(summary, use_container_width=True, hide_index=True)

    if selected_oc == "Tümü" or selected_pos == "Tümü":
        st.info("Etap detayını görmek için bir OC ve POS seç.")
        return

    item_row = filtered.iloc[0]
    item_id = int(item_row["item_id"])
    progress = get_operation_progress([item_id])
    st.markdown(f"### {selected_oc} / {selected_pos} etap ilerlemesi")
    if progress.empty:
        st.warning("Bu POS için kombinasyon henüz atanmadı. İlk işlem kaydında atanacaktır.")
    else:
        progress_table = progress[[
            "operation_order", "operation_name", "requested_qty", "completed_qty",
            "remaining_qty", "processed_qty", "fire_qty", "completion_pct",
        ]].rename(columns={
            "operation_order": "Sıra",
            "operation_name": "İşlem",
            "requested_qty": "İstenen",
            "completed_qty": "Sağlam İlerleyen",
            "remaining_qty": "Kalan",
            "processed_qty": "Toplam İşlenen",
            "fire_qty": "Fire",
            "completion_pct": "İlerleme %",
        })
        st.dataframe(progress_table, use_container_width=True, hide_index=True)

    st.markdown("### İşlem geçmişi")
    history = get_operation_history(selected_oc, selected_pos)
    if history.empty:
        st.info("Henüz işlem kaydı yok.")
    else:
        history_batches = history[
            [
                "batch_id",
                "calisma_saati",
                "operator_ismi",
            ]
        ].drop_duplicates("batch_id")
        total_history_hours = float(history_batches["calisma_saati"].sum())
        history_good = int(history["saglam_ilerleyen"].sum())
        history_area = float(history["saglam_alan_mm2"].sum())
        history_processed = int(history["islem_yapilan"].sum())
        history_fire = int(history["fire"].sum())

        h1, h2, h3, h4 = st.columns(4)
        h1.metric(
            "Geçmiş adet/saat",
            f"{history_good / total_history_hours:.2f}"
            if total_history_hours
            else "0.00",
        )
        h2.metric(
            "Geçmiş mm²/saat",
            f"{_format_mm2(history_area / total_history_hours)}"
            if total_history_hours
            else "0.00",
        )
        h3.metric(
            "Fire oranı",
            f"%{history_fire / history_processed * 100:.2f}"
            if history_processed
            else "%0.00",
        )
        h4.metric(
            "Toplam çalışma",
            f"{total_history_hours:.1f} saat",
        )

        st.dataframe(history, use_container_width=True, hide_index=True)
        st.download_button(
            "İşlem geçmişini CSV indir",
            data=history.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"OC_{selected_oc}_{selected_pos}_operasyon_gecmisi.csv",
            mime="text/csv",
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
    a1.metric("İstenen alan", f"{_format_m2(requested_area)} m²")
    a2.metric("Üretilen alan", f"{_format_m2(produced_area)} m²")
    a3.metric("Kalan alan", f"{_format_m2(remaining_area)} m²")

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



def rebuild_operation_completion_logs(
    conn: sqlite3.Connection,
    item_ids: list[int],
):
    """
    Operasyon kayıtları silindiğinde POS'un tam biten adet günlüklerini
    kalan operasyon geçmişinden kronolojik olarak yeniden oluşturur.
    """
    if not item_ids:
        return

    cur = conn.cursor()
    created_at = datetime.now().isoformat(timespec="seconds")

    for item_id in sorted({int(value) for value in item_ids}):
        cur.execute(
            """
            DELETE FROM production_output_logs
            WHERE item_id = ?
              AND source_type = 'operation_tracking'
            """,
            (item_id,),
        )

        cur.execute(
            """
            SELECT requested_qty, unit_area_mm2
            FROM production_output_items
            WHERE id = ?
            """,
            (item_id,),
        )
        item_row = cur.fetchone()
        if item_row is None:
            continue

        requested_qty = int(item_row[0])
        unit_area = float(item_row[1])

        cur.execute(
            """
            SELECT operation_name
            FROM pos_operation_plan
            WHERE item_id = ?
            ORDER BY operation_order
            """,
            (item_id,),
        )
        operations = [str(row[0]) for row in cur.fetchall()]
        if not operations:
            continue

        cur.execute(
            """
            SELECT
                b.id,
                b.tarih,
                b.operator_ismi,
                w.operation_name,
                w.good_qty
            FROM operation_work_logs w
            INNER JOIN operation_batches b ON b.id = w.batch_id
            WHERE w.item_id = ?
            ORDER BY b.tarih, b.id, w.id
            """,
            (item_id,),
        )
        rows = cur.fetchall()
        if not rows:
            continue

        totals = {operation: 0 for operation in operations}
        previous_completed = 0
        current_batch_id = None
        current_date = ""
        current_operator = ""
        batch_entries = []

        def apply_batch():
            nonlocal previous_completed
            if not batch_entries:
                return

            for operation_name, good_qty in batch_entries:
                if operation_name in totals:
                    totals[operation_name] += int(good_qty)

            completed = min(
                min(int(totals[operation]), requested_qty)
                for operation in operations
            )
            delta = max(completed - previous_completed, 0)

            if delta > 0:
                cur.execute(
                    """
                    INSERT INTO production_output_logs (
                        item_id,
                        production_date,
                        operator_name,
                        produced_qty,
                        produced_area_mm2,
                        note,
                        created_at,
                        source_type,
                        source_ref
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        item_id,
                        current_date,
                        current_operator,
                        delta,
                        delta * unit_area,
                        "Operasyon geçmişinden yeniden hesaplandı",
                        created_at,
                        "operation_tracking",
                        current_batch_id,
                    ),
                )
            previous_completed = completed

        for (
            batch_id,
            production_date,
            operator_name,
            operation_name,
            good_qty,
        ) in rows:
            if current_batch_id is None:
                current_batch_id = int(batch_id)
                current_date = str(production_date)
                current_operator = str(operator_name)

            if int(batch_id) != current_batch_id:
                apply_batch()
                batch_entries = []
                current_batch_id = int(batch_id)
                current_date = str(production_date)
                current_operator = str(operator_name)

            batch_entries.append(
                (str(operation_name), int(good_qty))
            )

        apply_batch()


def delete_operation_batch(batch_id: int) -> dict:
    """
    Seçilen operasyon kaydını ve bağlı işlem satırlarını siler.
    POS ilerlemeleri kalan kayıtlara göre yeniden hesaplanır.
    """
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()

        cur.execute(
            """
            SELECT DISTINCT item_id
            FROM operation_work_logs
            WHERE batch_id = ?
            """,
            (int(batch_id),),
        )
        item_ids = [int(row[0]) for row in cur.fetchall()]

        cur.execute(
            """
            SELECT COUNT(*)
            FROM operation_work_logs
            WHERE batch_id = ?
            """,
            (int(batch_id),),
        )
        deleted_detail_count = int(cur.fetchone()[0] or 0)

        cur.execute(
            """
            DELETE FROM operation_work_logs
            WHERE batch_id = ?
            """,
            (int(batch_id),),
        )
        cur.execute(
            """
            DELETE FROM operation_batches
            WHERE id = ?
            """,
            (int(batch_id),),
        )
        deleted_batch_count = int(cur.rowcount or 0)

        rebuild_operation_completion_logs(conn, item_ids)

        conn.commit()
        return {
            "deleted_batch_count": deleted_batch_count,
            "deleted_detail_count": deleted_detail_count,
            "affected_item_count": len(item_ids),
        }
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()



def get_operation_batch_summary() -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        result = pd.read_sql_query(
            """
            SELECT
                b.id AS batch_id,
                b.tarih,
                b.operator_ismi,
                b.calisma_tipi,
                b.calisma_saati,
                b.neden,
                GROUP_CONCAT(DISTINCT i.oc_no) AS OC,
                GROUP_CONCAT(DISTINCT i.pos) AS POS,
                GROUP_CONCAT(DISTINCT w.operation_name) AS islemler,
                SUM(w.processed_qty) AS islem_yapilan,
                SUM(w.fire_qty) AS fire,
                SUM(w.good_qty) AS saglam_ilerleyen,
                SUM(w.good_qty * i.unit_area_mm2) AS saglam_alan_mm2,
                SUM(
                    CASE
                        WHEN COALESCE(t.target_qty_per_hour, 0) > 0
                        THEN w.good_qty / t.target_qty_per_hour
                        ELSE 0
                    END
                ) AS target_qty_hours,
                SUM(
                    CASE
                        WHEN COALESCE(t.target_area_per_hour, 0) > 0
                        THEN (w.good_qty * i.unit_area_mm2)
                             / t.target_area_per_hour
                        ELSE 0
                    END
                ) AS target_area_hours,
                AVG(COALESCE(t.fire_limit_pct, 5)) AS fire_limit_pct,
                AVG(COALESCE(t.slow_limit_pct, 80)) AS slow_limit_pct,
                AVG(COALESCE(t.fast_limit_pct, 120)) AS fast_limit_pct,
                COUNT(*) AS islem_satiri,
                b.notlar,
                b.created_at
            FROM operation_batches b
            INNER JOIN operation_work_logs w ON w.batch_id = b.id
            INNER JOIN production_output_items i ON i.id = w.item_id
            LEFT JOIN operation_performance_targets t
                ON t.operation_name = w.operation_name
            GROUP BY
                b.id, b.tarih, b.operator_ismi, b.calisma_tipi,
                b.calisma_saati, b.neden, b.notlar, b.created_at
            ORDER BY b.tarih DESC, b.id DESC
            """,
            conn,
        )
    finally:
        conn.close()

    return enrich_operation_batch_performance(result)


def get_operation_batch_details(batch_id: int) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql_query(
            """
            SELECT
                w.id AS kayıt_no,
                i.oc_no AS OC,
                i.pos AS POS,
                w.operation_name AS İşlem,
                w.processed_qty AS İşlem_Yapılan,
                w.fire_qty AS Fire,
                w.good_qty AS Sağlam_İlerleyen,
                i.unit_area_mm2 AS Birim_Alan_mm2,
                w.good_qty * i.unit_area_mm2 AS Sağlam_Alan_mm2,
                CASE
                    WHEN w.processed_qty > 0
                    THEN ROUND(w.fire_qty * 100.0 / w.processed_qty, 2)
                    ELSE 0
                END AS Fire_Yüzde,
                COALESCE(t.fire_limit_pct, 5) AS Fire_Sınırı_Yüzde,
                CASE
                    WHEN w.processed_qty > 0
                         AND w.fire_qty * 100.0 / w.processed_qty
                             > COALESCE(t.fire_limit_pct, 5)
                    THEN 'Sınır Aşıldı'
                    ELSE 'Normal'
                END AS Fire_Durumu,
                w.fire_note AS Fire_Notu,
                w.created_at AS Kayıt_Zamanı
            FROM operation_work_logs w
            INNER JOIN production_output_items i ON i.id = w.item_id
            LEFT JOIN operation_performance_targets t
                ON t.operation_name = w.operation_name
            WHERE w.batch_id = ?
            ORDER BY i.oc_no, i.pos, w.operation_name
            """,
            conn,
            params=(int(batch_id),),
        )
    finally:
        conn.close()


def render_operation_day_end_section() -> bool:
    history = get_operation_history()
    if history.empty:
        return False

    history["tarih_dt"] = pd.to_datetime(history["tarih"], errors="coerce")
    history = history.dropna(subset=["tarih_dt"]).copy()
    if history.empty:
        return False

    available_dates = sorted(
        history["tarih_dt"].dt.date.unique(),
        reverse=True,
    )

    st.subheader("Operasyon Bazlı Gün Sonu")
    st.caption(
        "Yeni operasyon sistemiyle girilen lazer, abkant, kaynak, boya ve "
        "diğer işlem kayıtlarının günlük özetidir."
    )

    selected_date = st.date_input(
        "Operasyon kayıtları için kontrol tarihi",
        value=available_dates[0],
        min_value=min(available_dates),
        max_value=max(available_dates),
        key="operation_day_end_date",
    )

    day = history[
        history["tarih_dt"].dt.date == selected_date
    ].copy()

    if day.empty:
        st.warning("Seçilen tarihte operasyon kaydı bulunamadı.")
        return True

    batch_rows = day[
        [
            "batch_id",
            "operator_ismi",
            "calisma_tipi",
            "calisma_saati",
        ]
    ].drop_duplicates("batch_id")

    total_processed = int(day["islem_yapilan"].sum())
    total_good = int(day["saglam_ilerleyen"].sum())
    total_fire = int(day["fire"].sum())
    total_good_area = float(day["saglam_alan_mm2"].sum())
    total_hours = float(batch_rows["calisma_saati"].sum())
    productivity = total_good / total_hours if total_hours else 0.0
    area_productivity = (
        total_good_area / total_hours
        if total_hours
        else 0.0
    )
    day_fire_rate = (
        total_fire / total_processed * 100
        if total_processed
        else 0.0
    )

    day_entries = []
    for _, row in day.iterrows():
        day_entries.append(
            {
                "pos": row["POS"],
                "operation_name": row["operasyon"],
                "processed_qty": int(row["islem_yapilan"]),
                "fire_qty": int(row["fire"]),
                "good_qty": int(row["saglam_ilerleyen"]),
                "unit_area_mm2": float(row["unit_area_mm2"]),
            }
        )
    day_performance = calculate_entries_performance(
        day_entries,
        total_hours,
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Çalışan", int(day["operator_ismi"].nunique()))
    m2.metric("Günlük kayıt", int(day["batch_id"].nunique()))
    m3.metric("İşlem yapılan", f"{total_processed} adet")
    m4.metric("Sağlam ilerleyen", f"{total_good} adet")

    m5, m6, m7, m8 = st.columns(4)
    m5.metric(
        "Fire",
        f"{total_fire} adet",
        delta=f"%{day_fire_rate:.2f}",
        delta_color="inverse",
    )
    m6.metric("Toplam çalışma", f"{total_hours:.1f} saat")
    m7.metric("Adet / saat", f"{productivity:.2f}")
    m8.metric(
        "mm² / saat",
        f"{_format_mm2(area_productivity)}",
    )

    g1, g2, g3 = st.columns(3)
    g1.metric("Gidişat", day_performance["status"])
    g2.metric(
        "Verim skoru",
        (
            f"%{day_performance['score_pct']:.1f}"
            if day_performance["score_pct"] is not None
            else "Referans yok"
        ),
    )
    g3.metric("Çalışılan POS", int(day["POS"].nunique()))

    if day_performance["status"] == "Yavaş":
        st.warning(
            "Günün üretim gidişatı yavaş. Referans: "
            + day_performance["reference_source"]
        )
    elif day_performance["status"] == "Hızlı":
        st.success(
            "Günün üretim gidişatı hızlı. Referans: "
            + day_performance["reference_source"]
        )
    elif day_performance["status"] == "Normal":
        st.info(
            "Günün üretim gidişatı normal. Referans: "
            + day_performance["reference_source"]
        )

    for breach in day_performance["fire_breaches"]:
        st.error(
            f"{breach['pos']} / {breach['operation']}: "
            f"Fire %{breach['fire_rate_pct']:.1f}; "
            f"sınır %{breach['fire_limit_pct']:.1f}."
        )

    overtime = batch_rows[
        batch_rows["calisma_tipi"] == "Mesaili"
    ]
    if not overtime.empty:
        names = ", ".join(
            sorted(day[
                day["batch_id"].isin(overtime["batch_id"])
            ]["operator_ismi"].unique())
        )
        st.warning(f"Mesaili çalışma kaydı var: {names}")

    if total_fire > 0:
        st.error(f"Seçilen günde toplam {total_fire} adet operasyon firesi var.")
    else:
        st.success("Seçilen günde operasyon firesi bulunmuyor.")

    pos_summary = (
        day.groupby(["OC", "POS"], as_index=False)
        .agg(
            işlem_satırı=("operasyon", "count"),
            farklı_işlem=("operasyon", "nunique"),
            işlem_yapılan=("islem_yapilan", "sum"),
            sağlam_ilerleyen=("saglam_ilerleyen", "sum"),
            fire=("fire", "sum"),
        )
        .sort_values(["OC", "POS"])
    )

    operation_summary = (
        day.groupby("operasyon", as_index=False)
        .agg(
            kayıt=("batch_id", "nunique"),
            çalışan=("operator_ismi", "nunique"),
            işlem_yapılan=("islem_yapilan", "sum"),
            sağlam_ilerleyen=("saglam_ilerleyen", "sum"),
            fire=("fire", "sum"),
        )
        .sort_values("operasyon")
    )

    worker_activity = (
        day.groupby("operator_ismi", as_index=False)
        .agg(
            kayıt=("batch_id", "nunique"),
            POS=("POS", "nunique"),
            işlem=("operasyon", "nunique"),
            işlem_yapılan=("islem_yapilan", "sum"),
            sağlam_ilerleyen=("saglam_ilerleyen", "sum"),
            sağlam_alan_mm2=("saglam_alan_mm2", "sum"),
            fire=("fire", "sum"),
        )
    )
    worker_hours = (
        batch_rows.groupby("operator_ismi", as_index=False)
        .agg(çalışma_saati=("calisma_saati", "sum"))
    )
    worker_summary = worker_activity.merge(
        worker_hours,
        on="operator_ismi",
        how="left",
    )
    worker_summary["adet_saat"] = (
        worker_summary["sağlam_ilerleyen"]
        / worker_summary["çalışma_saati"].replace(0, 1)
    ).round(2)
    worker_summary["mm2_saat"] = (
        worker_summary["sağlam_alan_mm2"]
        / worker_summary["çalışma_saati"].replace(0, 1)
    ).round(2)
    worker_summary["fire_yuzde"] = (
        worker_summary["fire"]
        / worker_summary["işlem_yapılan"].replace(0, 1)
        * 100
    ).round(2)
    worker_summary = worker_summary.sort_values(
        "sağlam_ilerleyen",
        ascending=False,
    )

    detail = day[
        [
            "batch_id",
            "tarih",
            "operator_ismi",
            "calisma_tipi",
            "calisma_saati",
            "OC",
            "POS",
            "operasyon",
            "islem_yapilan",
            "fire",
            "saglam_ilerleyen",
            "saglam_alan_mm2",
            "fire_siniri_yuzde",
            "fire_notu",
            "notlar",
        ]
    ].copy()

    tabs = st.tabs([
        "POS Özeti",
        "İşlem Özeti",
        "İşçi Özeti",
        "Günlük Ayrıntı",
    ])

    with tabs[0]:
        st.dataframe(
            pos_summary,
            use_container_width=True,
            hide_index=True,
        )
        if not pos_summary.empty:
            st.bar_chart(
                pos_summary.set_index("POS")[
                    ["işlem_yapılan", "sağlam_ilerleyen", "fire"]
                ]
            )

    with tabs[1]:
        st.dataframe(
            operation_summary,
            use_container_width=True,
            hide_index=True,
        )
        if not operation_summary.empty:
            st.bar_chart(
                operation_summary.set_index("operasyon")[
                    ["işlem_yapılan", "sağlam_ilerleyen", "fire"]
                ]
            )

    with tabs[2]:
        st.dataframe(
            worker_summary,
            use_container_width=True,
            hide_index=True,
        )

    with tabs[3]:
        st.dataframe(
            detail,
            use_container_width=True,
            hide_index=True,
        )

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pos_summary.to_excel(
            writer,
            sheet_name="POS Özeti",
            index=False,
        )
        operation_summary.to_excel(
            writer,
            sheet_name="İşlem Özeti",
            index=False,
        )
        worker_summary.to_excel(
            writer,
            sheet_name="İşçi Özeti",
            index=False,
        )
        detail.to_excel(
            writer,
            sheet_name="Günlük Ayrıntı",
            index=False,
        )

    st.download_button(
        "Operasyon Gün Sonu Excel Raporunu İndir",
        data=output.getvalue(),
        file_name=f"operasyon_gun_sonu_{selected_date}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )

    return True



def kayitlar_page():
    st.subheader("Kayıtlar")
    st.caption(
        "Yeni operasyon bazlı kayıtlar ile önceki sistem kayıtları "
        "ayrı sekmelerde gösterilir."
    )

    operation_batches = get_operation_batch_summary()
    operation_tab, legacy_tab = st.tabs(
        ["Operasyon Kayıtları", "Eski Sistem Kayıtları"]
    )

    with operation_tab:
        if operation_batches.empty:
            render_empty_state("Henüz operasyon kaydı yok", "İlk çalışan kaydı tamamlandığında ayrıntılar, verim ve silme seçenekleri burada görünür.", "▤")
        else:
            display_batches = operation_batches.rename(
                columns={
                    "batch_id": "Kayıt No",
                    "tarih": "Tarih",
                    "operator_ismi": "İşçi",
                    "calisma_tipi": "Çalışma Tipi",
                    "calisma_saati": "Çalışma Saati",
                    "neden": "Neden",
                    "islemler": "İşlemler",
                    "islem_yapilan": "İşlem Yapılan",
                    "fire": "Fire",
                    "saglam_ilerleyen": "Sağlam İlerleyen",
                    "saglam_alan_mm2": "Sağlam Alan mm²",
                    "adet_saat": "Adet/Saat",
                    "mm2_saat": "mm²/Saat",
                    "fire_yuzde": "Fire %",
                    "fire_durumu": "Fire Durumu",
                    "gidisat": "Gidişat",
                    "verim_skoru_yuzde": "Verim Skoru %",
                    "verim_referansi": "Verim Referansı",
                    "islem_satiri": "İşlem Satırı",
                    "notlar": "Not",
                    "created_at": "Kayıt Zamanı",
                }
            )
            st.dataframe(
                display_batches,
                use_container_width=True,
                hide_index=True,
            )

            labels = operation_batches.apply(
                lambda row: (
                    f"#{int(row['batch_id'])} · {row['tarih']} · "
                    f"{row['operator_ismi']} · OC {row['OC']} · "
                    f"{row['POS']}"
                ),
                axis=1,
            ).tolist()
            selected_label = st.selectbox(
                "Ayrıntısını görmek istediğin operasyon kaydı",
                labels,
                key="operation_record_select",
            )
            selected_index = labels.index(selected_label)
            selected_batch = operation_batches.iloc[selected_index]
            selected_batch_id = int(selected_batch["batch_id"])

            details = get_operation_batch_details(selected_batch_id)

            k1, k2, k3, k4 = st.columns(4)
            k1.metric(
                "İşlem yapılan",
                f"{int(selected_batch['islem_yapilan'])} adet",
            )
            k2.metric(
                "Sağlam ilerleyen",
                f"{int(selected_batch['saglam_ilerleyen'])} adet",
            )
            k3.metric(
                "Fire",
                f"{int(selected_batch['fire'])} adet",
                delta=f"%{float(selected_batch['fire_yuzde']):.2f}",
                delta_color="inverse",
            )
            k4.metric(
                "Çalışma",
                f"{float(selected_batch['calisma_saati']):.1f} saat",
            )

            v1, v2, v3, v4 = st.columns(4)
            v1.metric(
                "Adet / saat",
                f"{float(selected_batch['adet_saat']):.2f}",
            )
            v2.metric(
                "mm² / saat",
                f"{_format_mm2(selected_batch['mm2_saat'])}",
            )
            v3.metric(
                "Gidişat",
                str(selected_batch["gidisat"]),
                delta=(
                    f"%{float(selected_batch['verim_skoru_yuzde']):.1f}"
                    if pd.notna(selected_batch["verim_skoru_yuzde"])
                    else str(selected_batch["verim_referansi"])
                ),
            )
            v4.metric(
                "Fire durumu",
                str(selected_batch["fire_durumu"]),
            )

            if selected_batch["gidisat"] == "Yavaş":
                st.warning("Bu kaydın gidişatı referansa göre yavaş.")
            elif selected_batch["gidisat"] == "Hızlı":
                st.success("Bu kaydın gidişatı referansa göre hızlı.")

            if selected_batch["fire_durumu"] == "Sınır Aşıldı":
                st.error(
                    f"Fire oranı %{float(selected_batch['fire_yuzde']):.2f}; "
                    f"ortalama sınır %{float(selected_batch['fire_limit_pct']):.2f}."
                )

            st.write("İşlem ayrıntıları")
            st.dataframe(
                details,
                use_container_width=True,
                hide_index=True,
            )

            st.divider()
            st.subheader("Operasyon Kaydını Sil")
            st.warning(
                f"#{selected_batch_id} numaralı operasyon kaydı silinirse "
                "bu kayıttaki işlem adetleri, fireler ve POS ilerlemeleri "
                "geri alınacaktır."
            )

            delete_operation_confirm = st.checkbox(
                "Bu operasyon kaydını kalıcı olarak silmek istediğimi onaylıyorum",
                key=f"delete_operation_confirm_{selected_batch_id}",
            )

            if st.button(
                "Seçili Operasyon Kaydını Sil",
                type="primary",
                disabled=not delete_operation_confirm,
                key=f"delete_operation_button_{selected_batch_id}",
            ):
                result = delete_operation_batch(selected_batch_id)
                st.success(
                    f"#{selected_batch_id} numaralı operasyon kaydı silindi. "
                    f"{result['deleted_detail_count']} işlem satırı kaldırıldı ve "
                    f"{result['affected_item_count']} POS yeniden hesaplandı."
                )
                st.rerun()

            st.divider()
            csv = operation_batches.to_csv(
                index=False,
            ).encode("utf-8-sig")
            st.download_button(
                "Operasyon Kayıtlarını CSV İndir",
                data=csv,
                file_name="operasyon_kayitlari.csv",
                mime="text/csv",
            )

    with legacy_tab:
        conn = sqlite3.connect(DB_PATH)
        sessions = pd.read_sql_query(
            "SELECT * FROM work_sessions ORDER BY id DESC",
            conn,
        )

        if sessions.empty:
            st.info("Eski sistemde kayıt bulunmuyor.")
            conn.close()
        else:
            sessions["durum"] = sessions.apply(
                lambda row: production_status(
                    int(row["siparis_adedi"]),
                    int(row["saglam_tamamlanan"]),
                ),
                axis=1,
            )
            sessions["kalan"] = (
                sessions["siparis_adedi"]
                - sessions["saglam_tamamlanan"]
            ).clip(lower=0)

            st.dataframe(
                _rename_oc_column(sessions),
                use_container_width=True,
            )

            selected_id = st.selectbox(
                "Eski kayıt ayrıntısı",
                sessions["id"].tolist(),
                key="legacy_record_select",
            )
            details = pd.read_sql_query(
                """
                SELECT
                    urun_no, operasyon_sirasi, operasyon_adi,
                    yapildi, fire_var, fire_operasyonu, fire_notu
                FROM operation_entries
                WHERE session_id = ?
                ORDER BY urun_no, operasyon_sirasi
                """,
                conn,
                params=(int(selected_id),),
            )
            conn.close()

            st.write("Operasyon detayları")
            st.dataframe(details, use_container_width=True)

            selected_session = sessions[
                sessions["id"] == selected_id
            ].iloc[0]
            render_operation_progress(
                details,
                selected_session,
                int(selected_id),
            )

            st.divider()
            st.subheader("Güncelleme Geçmişi")
            history = get_update_history(int(selected_id))
            if history.empty:
                st.info("Bu kayıt henüz sonradan güncellenmemiş.")
            else:
                st.dataframe(history, use_container_width=True)

            st.divider()
            st.subheader("Eski Kayıt Silme")
            st.warning(
                f"{selected_id} numaralı eski kayıt ve bağlı ayrıntıları "
                "kalıcı olarak silinecek."
            )
            silme_onayi = st.checkbox(
                "Bu eski kaydı silmek istediğimi onaylıyorum",
                key=f"delete_confirm_{selected_id}",
            )
            if st.button(
                "Seçili Eski Kaydı Sil",
                type="primary",
                disabled=not silme_onayi,
                key=f"delete_button_{selected_id}",
            ):
                delete_session(int(selected_id))
                st.success(f"{selected_id} numaralı kayıt silindi.")
                st.rerun()

            csv = _rename_oc_column(sessions).to_csv(
                index=False,
            ).encode("utf-8-sig")
            st.download_button(
                "Eski Kayıtları CSV İndir",
                csv,
                "eski_uretim_kayitlari.csv",
                "text/csv",
            )



def grafikler_page():
    st.subheader("Üretim Analizleri")
    st.caption(
        "Günlük üretim eğrisi, hedef-gerçekleşen, adet/saat, mm²/saat, fire ve iş dağılımını birlikte gösterir."
    )

    history = get_operation_history()

    if history.empty:
        render_empty_state(
            "Grafik oluşturacak operasyon kaydı yok",
            "İlk operasyon kaydı oluşturulduğunda günlük üretim, verim ve fire grafikleri burada otomatik oluşur.",
            "▥",
        )
    else:
        history = history.copy()
        history["tarih_dt"] = pd.to_datetime(history["tarih"], errors="coerce")
        history = history.dropna(subset=["tarih_dt"])

        batch_hours = (
            history[["batch_id", "tarih_dt", "operator_ismi", "calisma_saati"]]
            .drop_duplicates("batch_id")
        )
        daily_hours = batch_hours.groupby("tarih_dt", as_index=False).agg(
            çalışma_saati=("calisma_saati", "sum")
        )
        daily = history.groupby("tarih_dt", as_index=False).agg(
            işlem_yapılan=("islem_yapilan", "sum"),
            sağlam_ilerleyen=("saglam_ilerleyen", "sum"),
            fire=("fire", "sum"),
            sağlam_alan_mm2=("saglam_alan_mm2", "sum"),
        ).merge(daily_hours, on="tarih_dt", how="left")
        daily["adet_saat"] = (
            daily["sağlam_ilerleyen"] / daily["çalışma_saati"].replace(0, 1)
        ).round(2)
        daily["mm2_saat"] = (
            daily["sağlam_alan_mm2"] / daily["çalışma_saati"].replace(0, 1)
        ).round(2)
        daily["fire_yüzde"] = (
            daily["fire"] / daily["işlem_yapılan"].replace(0, 1) * 100
        ).round(2)
        daily = daily.sort_values("tarih_dt")

        batch_processed = history.groupby("batch_id", as_index=False).agg(
            batch_toplam=("islem_yapilan", "sum")
        )
        row_perf = history.merge(batch_processed, on="batch_id", how="left")
        row_perf["ayrılan_saat"] = (
            row_perf["calisma_saati"]
            * row_perf["islem_yapilan"]
            / row_perf["batch_toplam"].replace(0, 1)
        )
        operation_perf = row_perf.groupby("operasyon", as_index=False).agg(
            sağlam_ilerleyen=("saglam_ilerleyen", "sum"),
            sağlam_alan_mm2=("saglam_alan_mm2", "sum"),
            işlem_yapılan=("islem_yapilan", "sum"),
            fire=("fire", "sum"),
            ayrılan_saat=("ayrılan_saat", "sum"),
            hedef_adet_saat=("hedef_adet_saat", "max"),
            hedef_mm2_saat=("hedef_mm2_saat", "max"),
            fire_sınırı=("fire_siniri_yuzde", "max"),
        )
        operation_perf["gerçek_adet_saat"] = (
            operation_perf["sağlam_ilerleyen"]
            / operation_perf["ayrılan_saat"].replace(0, 1)
        ).round(2)
        operation_perf["gerçek_mm2_saat"] = (
            operation_perf["sağlam_alan_mm2"]
            / operation_perf["ayrılan_saat"].replace(0, 1)
        ).round(2)
        operation_perf["fire_yüzde"] = (
            operation_perf["fire"]
            / operation_perf["işlem_yapılan"].replace(0, 1)
            * 100
        ).round(2)

        numeric_columns = [
            "sağlam_ilerleyen",
            "sağlam_alan_mm2",
            "işlem_yapılan",
            "fire",
            "ayrılan_saat",
            "hedef_adet_saat",
            "hedef_mm2_saat",
            "fire_sınırı",
            "gerçek_adet_saat",
            "gerçek_mm2_saat",
            "fire_yüzde",
        ]
        for column_name in numeric_columns:
            if column_name in operation_perf.columns:
                operation_perf[column_name] = pd.to_numeric(
                    operation_perf[column_name],
                    errors="coerce",
                )

        worker = history.groupby("operator_ismi", as_index=False).agg(
            sağlam_ilerleyen=("saglam_ilerleyen", "sum"),
            sağlam_alan_mm2=("saglam_alan_mm2", "sum"),
            fire=("fire", "sum"),
            işlem=("operasyon", "nunique"),
            POS=("POS", "nunique"),
        ).sort_values("sağlam_ilerleyen", ascending=False)

        overview = get_operation_overview()

        latest = daily.iloc[-1]
        prior = daily.iloc[:-1]
        prior_qty = float(prior["adet_saat"].mean()) if not prior.empty else float(latest["adet_saat"])
        change_pct = (
            (float(latest["adet_saat"]) - prior_qty) / prior_qty * 100
            if prior_qty > 0 else 0.0
        )
        worst_fire = operation_perf.sort_values("fire_yüzde", ascending=False).iloc[0]
        best_op = operation_perf.sort_values("gerçek_adet_saat", ascending=False).iloc[0]
        slow_pos_text = "Veri yok"
        if not overview.empty:
            slow_pos = overview.sort_values("completion_pct").iloc[0]
            slow_pos_text = f"{slow_pos['pos']} · %{float(slow_pos['completion_pct']):.1f}"

        trend_icon = "↗" if change_pct >= 0 else "↘"
        trend_word = "arttı" if change_pct >= 0 else "azaldı"
        render_insight_cards([
            {
                "icon": trend_icon,
                "title": "Günlük verim eğilimi",
                "copy": f"Son gün adet/saat değeri önceki ortalamaya göre %{abs(change_pct):.1f} {trend_word}.",
            },
            {
                "icon": "⚡",
                "title": "En yüksek adet/saat",
                "copy": f"{best_op['operasyon']} · {float(best_op['gerçek_adet_saat']):.2f} adet/saat.",
            },
            {
                "icon": "!",
                "title": "En yüksek fire",
                "copy": f"{worst_fire['operasyon']} · %{float(worst_fire['fire_yüzde']):.2f} fire.",
            },
            {
                "icon": "◎",
                "title": "En yavaş ilerleyen POS",
                "copy": slow_pos_text,
            },
        ])

        t1, t2, t3, t4, t5 = st.tabs([
            "Günlük Eğilim",
            "Hedef / Gerçekleşen",
            "Fire Analizi",
            "POS Gidişatı",
            "İş Dağılımı",
        ])

        with t1:
            st.write("Günlük sağlam üretim")
            st.line_chart(
                daily.set_index("tarih_dt")[["sağlam_ilerleyen", "işlem_yapılan"]]
            )
            c1, c2 = st.columns(2)
            with c1:
                st.write("Adet / saat")
                st.line_chart(daily.set_index("tarih_dt")[["adet_saat"]])
            with c2:
                st.write("mm² / saat")
                st.line_chart(daily.set_index("tarih_dt")[["mm2_saat"]])
            if change_pct < -10:
                st.warning(
                    f"Son gün adet/saat değeri önceki günlerin ortalamasının %{abs(change_pct):.1f} altında."
                )
            elif change_pct > 10:
                st.success(
                    f"Son gün adet/saat değeri önceki günlerin ortalamasının %{change_pct:.1f} üzerinde."
                )
            else:
                st.info("Günlük verim son dönem ortalamasına yakın ilerliyor.")

        with t2:
            qty_compare = operation_perf[
                ["operasyon", "gerçek_adet_saat", "hedef_adet_saat"]
            ].set_index("operasyon")
            area_compare = operation_perf[
                ["operasyon", "gerçek_mm2_saat", "hedef_mm2_saat"]
            ].set_index("operasyon")
            st.write("Operasyon bazlı adet/saat hedefi")
            st.bar_chart(qty_compare)
            st.write("Operasyon bazlı mm²/saat hedefi")
            st.bar_chart(area_compare)
            actual_qty_rate = pd.to_numeric(
                operation_perf["gerçek_adet_saat"],
                errors="coerce",
            )
            target_qty_rate = pd.to_numeric(
                operation_perf["hedef_adet_saat"],
                errors="coerce",
            ).mask(
                lambda values: values <= 0
            )

            operation_perf["hedef_karşılama_%"] = (
                actual_qty_rate
                .div(target_qty_rate)
                .mul(100)
                .round(1)
            )
            st.dataframe(
                operation_perf[
                    [
                        "operasyon", "gerçek_adet_saat", "hedef_adet_saat",
                        "gerçek_mm2_saat", "hedef_mm2_saat", "hedef_karşılama_%",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )
            below = operation_perf[
                operation_perf["hedef_karşılama_%"].fillna(100) < 80
            ]
            missing_targets = operation_perf[
                operation_perf["hedef_adet_saat"].fillna(0) <= 0
            ]

            if not below.empty:
                st.warning(
                    "Hedefin %80 altında kalan işlemler: "
                    + ", ".join(below["operasyon"].astype(str))
                )
            else:
                st.success(
                    "Tanımlı hedeflerin altında kritik bir operasyon görünmüyor."
                )

            if not missing_targets.empty:
                st.info(
                    "Adet/saat hedefi henüz girilmemiş işlemler: "
                    + ", ".join(
                        missing_targets["operasyon"].astype(str)
                    )
                    + ". Yönetici, Verim ve Fire Hedefleri sayfasından "
                    "bu değerleri tanımlayabilir."
                )

        with t3:
            fire_chart = operation_perf[
                ["operasyon", "fire_yüzde", "fire_sınırı"]
            ].set_index("operasyon")
            st.bar_chart(fire_chart)
            st.dataframe(
                operation_perf[
                    ["operasyon", "işlem_yapılan", "fire", "fire_yüzde", "fire_sınırı"]
                ],
                use_container_width=True,
                hide_index=True,
            )
            breached = operation_perf[
                operation_perf["fire_yüzde"] > operation_perf["fire_sınırı"]
            ]
            if breached.empty:
                st.success("Bütün operasyonların fire oranı tanımlı sınırlar içinde.")
            else:
                st.error(
                    "Fire sınırını aşan işlemler: "
                    + ", ".join(breached["operasyon"].astype(str))
                )

        with t4:
            if overview.empty:
                render_empty_state(
                    "POS ilerleme verisi yok",
                    "Kombinasyon planları ve operasyon kayıtları oluştuğunda POS bazlı gidişat burada görünür.",
                    "◎",
                )
            else:
                pos_chart = overview.sort_values("completion_pct")[
                    ["pos", "completion_pct"]
                ].set_index("pos")
                st.bar_chart(pos_chart)
                st.dataframe(
                    overview.sort_values("completion_pct")[[
                        "oc_no", "pos", "completion_pct", "remaining_qty"
                    ]],
                    use_container_width=True,
                    hide_index=True,
                )
                near = overview[
                    (overview["completion_pct"] >= 85)
                    & (overview["completion_pct"] < 100)
                ]
                if not near.empty:
                    st.info(
                        "Tamamlanmaya yakın POS'lar: "
                        + ", ".join(near["pos"].astype(str))
                    )

        with t5:
            st.write("Çalışan bazlı sağlam ilerleme")
            st.bar_chart(worker.set_index("operator_ismi")[["sağlam_ilerleyen"]])
            st.dataframe(worker, use_container_width=True, hide_index=True)
            if not worker.empty:
                top_worker = worker.iloc[0]
                st.info(
                    f"En yüksek sağlam ilerleme: {top_worker['operator_ismi']} · "
                    f"{int(top_worker['sağlam_ilerleyen'])} işlem adedi."
                )

    # Eski grafik özelliğini kaybetmemek için ayrı bölümde koru.
    conn = sqlite3.connect(DB_PATH)
    legacy_sessions = pd.read_sql_query(
        "SELECT * FROM work_sessions ORDER BY id DESC",
        conn,
    )
    if not legacy_sessions.empty:
        with st.expander("Eski sistem kombinasyon grafiklerini göster", expanded=False):
            labels = legacy_sessions.apply(
                lambda r: f"#{int(r['id'])} - {r['tarih']} - {r['operator_ismi']} - {r['pos']} - {r['kombinasyon_adi']}",
                axis=1,
            ).tolist()
            selected_label = st.selectbox(
                "Eski kaydı seç",
                labels,
                key="legacy_graph_record",
            )
            selected_index = labels.index(selected_label)
            selected_session = legacy_sessions.iloc[selected_index]
            selected_id = int(selected_session["id"])
            details = pd.read_sql_query(
                "SELECT urun_no, operasyon_sirasi, operasyon_adi, yapildi, fire_var, fire_operasyonu, fire_notu "
                "FROM operation_entries WHERE session_id = ? ORDER BY urun_no, operasyon_sirasi",
                conn,
                params=(selected_id,),
            )
            if not details.empty:
                total_units = int(selected_session.get("uretim_yuku", 0) or 0)
                summary = (
                    details.groupby("operasyon_adi", as_index=False)["yapildi"]
                    .sum().rename(columns={"yapildi": "Yapılan"})
                )
                summary["Kalan"] = (total_units - summary["Yapılan"]).clip(lower=0)
                st.dataframe(summary, use_container_width=True, hide_index=True)
                st.bar_chart(summary.set_index("operasyon_adi")[["Yapılan", "Kalan"]])
            else:
                st.info("Bu eski kayıtta operasyon ayrıntısı bulunmuyor.")
    conn.close()


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
    operation_history = get_operation_history()
    if not operation_history.empty:
        st.markdown("### Operasyon bazlı güncel kayıtlar")
        batch_summary = (
            operation_history.drop_duplicates("batch_id")
            [["batch_id", "tarih", "operator_ismi", "calisma_tipi", "calisma_saati"]]
        )
        n1, n2, n3, n4 = st.columns(4)
        n1.metric("Günlük kayıt", int(operation_history["batch_id"].nunique()))
        n2.metric("İşlem yapılan", int(operation_history["islem_yapilan"].sum()))
        n3.metric("Sağlam ilerleyen", int(operation_history["saglam_ilerleyen"].sum()))
        n4.metric("Fire", int(operation_history["fire"].sum()))

        operation_worker_summary = (
            operation_history.groupby("operator_ismi", as_index=False)
            .agg(
                islem_satiri=("batch_id", "count"),
                islem_yapilan=("islem_yapilan", "sum"),
                saglam_ilerleyen=("saglam_ilerleyen", "sum"),
                fire=("fire", "sum"),
            )
        )
        hours_summary = (
            batch_summary.groupby("operator_ismi", as_index=False)
            .agg(calisma_saati=("calisma_saati", "sum"))
        )
        operation_worker_summary = operation_worker_summary.merge(
            hours_summary, on="operator_ismi", how="left"
        ).rename(columns={
            "operator_ismi": "İşçi",
            "islem_satiri": "İşlem Satırı",
            "islem_yapilan": "İşlem Yapılan",
            "saglam_ilerleyen": "Sağlam İlerleyen",
            "fire": "Fire",
            "calisma_saati": "Çalışma Saati",
        })
        st.dataframe(operation_worker_summary, use_container_width=True, hide_index=True)
        st.caption("Ayrıntılı POS ve etap ilerlemesi için Operasyon Takibi sayfasını kullan.")
        st.divider()

    if sessions.empty:
        if operation_history.empty:
            st.info("Yönetici paneli için henüz kayıt yok.")
        else:
            st.info("Eski kayıt sisteminde kayıt yok; yeni operasyon kayıtları yukarıda gösteriliyor.")
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
        "Seçilen günün operasyonlarını, çalışanlarını, firelerini ve "
        "çalışma sürelerini tek ekranda gösterir."
    )

    has_operation_records = render_operation_day_end_section()

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
        if not has_operation_records:
            st.info("Gün sonu kontrolü için henüz kayıt yok.")
        return

    if has_operation_records:
        st.divider()
        st.subheader("Eski Sistem Gün Sonu Kayıtları")
        st.caption(
            "Bu bölüm yalnızca önceki sürümlerde work_sessions tablosuna "
            "kaydedilmiş kayıtları gösterir."
        )

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
    st.caption(
        "Yeni operasyon sisteminin güncel üretim durumunu gösterir."
    )

    overview = get_operation_overview()
    operation_history = get_operation_history()

    if overview.empty and operation_history.empty:
        render_empty_state("Özet oluşturacak veri yok", "Üretim çıktısı ve ilk operasyon kayıtları eklendiğinde genel durum burada otomatik hazırlanır.", "◉")
        return

    if not overview.empty:
        total_requested = int(overview["requested_qty"].sum())
        total_completed = int(overview["produced_qty"].sum())
        total_in_production = int(
            overview["in_production_qty"].sum()
        )
        total_fire = int(overview["operation_fire_qty"].sum())
        total_remaining = max(
            total_requested - total_completed,
            0,
        )

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Toplam istenen", total_requested)
        c2.metric("Tam biten", total_completed)
        c3.metric("Üretimde olan", total_in_production)
        c4.metric("Kalan", total_remaining)

        c5, c6, c7, c8 = st.columns(4)
        c5.metric("Operasyon firesi", total_fire)
        c6.metric("OC sayısı", int(overview["oc_no"].nunique()))
        c7.metric("POS sayısı", int(overview["item_id"].nunique()))
        completion_rate = (
            total_completed / total_requested * 100
            if total_requested
            else 0.0
        )
        c8.metric("Tamamlanma", f"%{completion_rate:.1f}")

        summary_table = overview[
            [
                "oc_no",
                "project_name",
                "pos",
                "requested_qty",
                "produced_qty",
                "in_production_qty",
                "remaining_qty",
                "operation_fire_qty",
                "completion_pct",
            ]
        ].rename(
            columns={
                "oc_no": "OC",
                "project_name": "Proje",
                "pos": "POS",
                "requested_qty": "İstenen",
                "produced_qty": "Tam Biten",
                "in_production_qty": "Üretimde",
                "remaining_qty": "Kalan",
                "operation_fire_qty": "Fire",
                "completion_pct": "Tamamlanma %",
            }
        )

        st.write("POS bazlı güncel durum")
        st.dataframe(
            summary_table,
            use_container_width=True,
            hide_index=True,
        )

    if not operation_history.empty:
        st.divider()
        st.subheader("Operasyon Kayıt Özeti")

        batch_rows = operation_history[
            [
                "batch_id",
                "operator_ismi",
                "calisma_saati",
            ]
        ].drop_duplicates("batch_id")

        h1, h2, h3, h4 = st.columns(4)
        h1.metric(
            "Operasyon kaydı",
            int(operation_history["batch_id"].nunique()),
        )
        h2.metric(
            "İşlem yapılan",
            int(operation_history["islem_yapilan"].sum()),
        )
        h3.metric(
            "Sağlam ilerleyen",
            int(operation_history["saglam_ilerleyen"].sum()),
        )
        h4.metric(
            "Toplam çalışma",
            f"{float(batch_rows['calisma_saati'].sum()):.1f} saat",
        )

        by_operation = (
            operation_history.groupby(
                "operasyon",
                as_index=False,
            )
            .agg(
                işlem_yapılan=("islem_yapilan", "sum"),
                sağlam_ilerleyen=("saglam_ilerleyen", "sum"),
                fire=("fire", "sum"),
            )
        )
        st.dataframe(
            by_operation,
            use_container_width=True,
            hide_index=True,
        )
        if not by_operation.empty:
            st.bar_chart(
                by_operation.set_index("operasyon")[
                    ["işlem_yapılan", "sağlam_ilerleyen", "fire"]
                ]
            )

    conn = sqlite3.connect(DB_PATH)
    legacy_sessions = pd.read_sql_query(
        "SELECT * FROM work_sessions",
        conn,
    )
    conn.close()

    if not legacy_sessions.empty:
        with st.expander(
            "Eski sistem özetini göster",
            expanded=False,
        ):
            col1, col2, col3, col4 = st.columns(4)
            col1.metric(
                "Eski toplam sipariş",
                int(legacy_sessions["siparis_adedi"].sum()),
            )
            col2.metric(
                "Eski sağlam tamamlanan",
                int(legacy_sessions["saglam_tamamlanan"].sum()),
            )
            col3.metric(
                "Eski fire",
                int(legacy_sessions["fire_adedi"].sum()),
            )
            col4.metric(
                "Eski üretim yükü",
                int(legacy_sessions["uretim_yuku"].sum()),
            )



if __name__ == "__main__":
    main()
