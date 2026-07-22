import hmac
import os
import re
import sqlite3

try:
    import libsql
except ImportError:
    libsql = None
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


def _read_secret(name: str) -> str:
    """Streamlit Secrets veya ortam değişkeninden güvenli ayar okur."""
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = ""
    if not value:
        value = os.getenv(name, "")
    return str(value or "").strip()


def remote_database_configured() -> bool:
    return bool(
        _read_secret("TURSO_DATABASE_URL")
        and _read_secret("TURSO_AUTH_TOKEN")
    )


def running_on_streamlit_cloud() -> bool:
    """Community Cloud çalışma dizinini güvenli biçimde ayırt eder."""
    return str(BASE_DIR).replace("\\", "/").startswith("/mount/src/")


def get_db_connection():
    """
    Her veritabanı işlemi için bağımsız bir bağlantı açar.

    Ortak bağlantı/kilit kullanılmaz. Streamlit yeniden çalışmaları birbiriyle
    çakışsa bile bir ekran diğerini bekletmez. Veriler aynı Turso veritabanında
    kalıcı olarak tutulmaya devam eder.
    """
    database_url = _read_secret("TURSO_DATABASE_URL")
    auth_token = _read_secret("TURSO_AUTH_TOKEN")

    if database_url and auth_token:
        if libsql is None:
            raise RuntimeError(
                "Kalıcı veritabanı sürücüsü kurulu değil. "
                "requirements.txt dosyasına libsql==0.1.11 eklenmelidir."
            )
        return libsql.connect(
            database=database_url,
            auth_token=auth_token,
        )

    if running_on_streamlit_cloud():
        raise RuntimeError(
            "KALICI VERİTABANI BAĞLANTISI YOK. Veri kaybını önlemek için "
            "uygulama kayıt işlemlerini durdurdu. Streamlit Secrets bölümüne "
            "TURSO_DATABASE_URL ve TURSO_AUTH_TOKEN ekleyin."
        )

    return sqlite3.connect(DB_PATH)


def _clear_data_caches():
    """Yalnızca veritabanı okuma önbelleklerini temizler."""
    cached_readers = (
        "get_update_history",
        "get_production_output_summary",
        "get_production_output_history",
        "get_performance_targets",
        "get_historical_efficiency_reference",
        "get_worker_competencies",
        "get_competency_table",
        "get_item_operation_plan",
        "get_operation_progress",
        "get_operation_history",
        "get_overtime_summary",
        "get_operation_batch_summary",
        "get_operation_batch_details",
        "get_operation_overview",
        "get_other_work_logs",
        "get_daily_m2_target",
    )
    for function_name in cached_readers:
        function = globals().get(function_name)
        clear_method = getattr(function, "clear", None)
        if callable(clear_method):
            try:
                clear_method()
            except Exception:
                pass

def build_full_database_backup_excel() -> bytes:
    """Tüm uygulama tablolarını tek Excel yedeğinde toplar."""
    conn = get_db_connection()
    try:
        table_rows = conn.execute(
            "SELECT name FROM sqlite_master "
            "WHERE type = 'table' AND name NOT LIKE 'sqlite_%' "
            "ORDER BY name"
        ).fetchall()
        table_names = [str(row[0]) for row in table_rows]

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_rows = []
            for table_name in table_names:
                safe_table = table_name.replace('"', '""')
                frame = pd.read_sql_query(
                    f'SELECT * FROM "{safe_table}"',
                    conn,
                )
                sheet_name = table_name[:31]
                frame.to_excel(writer, sheet_name=sheet_name, index=False)
                summary_rows.append(
                    {"tablo": table_name, "satir_sayisi": int(len(frame))}
                )

            pd.DataFrame(summary_rows).to_excel(
                writer,
                sheet_name="YEDEK_OZETI",
                index=False,
            )
        return output.getvalue()
    finally:
        conn.close()


def render_manager_backup_download():
    """Yedeği yalnızca kullanıcı istediğinde hazırlar; her ekran yenilemesinde sorgu çalıştırmaz."""
    st.markdown("**Veri Güvenliği**")

    if st.button(
        "Excel yedeğini hazırla",
        use_container_width=True,
        key="prepare_full_database_backup",
    ):
        try:
            with st.spinner("Veritabanı yedeği hazırlanıyor..."):
                st.session_state["full_database_backup_bytes"] = (
                    build_full_database_backup_excel()
                )
                st.session_state["full_database_backup_time"] = (
                    datetime.now().strftime("%Y%m%d_%H%M%S")
                )
        except Exception as exc:
            st.error(f"Yedek hazırlanamadı: {exc}")

    backup_bytes = st.session_state.get("full_database_backup_bytes")
    if backup_bytes:
        backup_time = st.session_state.get(
            "full_database_backup_time",
            datetime.now().strftime("%Y%m%d_%H%M%S"),
        )
        st.download_button(
            "Hazırlanan yedeği indir",
            data=backup_bytes,
            file_name=f"uretim_takip_yedek_{backup_time}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
            key="download_full_database_backup",
        )

    st.caption(
        "Yedek yalnızca bu düğmeye bastığında hazırlanır; böylece uygulama "
        "her tıklamada tüm veritabanını tekrar okumaz."
    )


def render_database_status():
    if remote_database_configured():
        st.success(
            "☁️ Kalıcı bulut veritabanı bağlı. Kod güncellense veya uygulama "
            "yeniden başlasa da kayıtlar korunur."
        )
    else:
        st.info(
            "💻 Yerel SQLite kullanılıyor. Bu mod yalnızca bilgisayarda "
            "çalıştırma içindir."
        )


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
@st.cache_resource(show_spinner=False)
def init_db():
    conn = get_db_connection()
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
            operation_hours REAL NOT NULL DEFAULT 0,
            laser_plate_qty INTEGER NOT NULL DEFAULT 0,
            material_type TEXT NOT NULL DEFAULT '',
            thickness_mm REAL NOT NULL DEFAULT 0,
            laser_long_edge_qty INTEGER NOT NULL DEFAULT 0,
            laser_short_edge_qty INTEGER NOT NULL DEFAULT 0,
            laser_long_per_team INTEGER NOT NULL DEFAULT 2,
            laser_short_per_team INTEGER NOT NULL DEFAULT 2,
            abkant_work_mode TEXT NOT NULL DEFAULT '',
            abkant_coworker TEXT NOT NULL DEFAULT '',
            abkant_team_qty INTEGER NOT NULL DEFAULT 0,
            abkant_teams_per_piece INTEGER NOT NULL DEFAULT 1,
            abkant_long_bend_qty INTEGER NOT NULL DEFAULT 0,
            abkant_short_bend_qty INTEGER NOT NULL DEFAULT 0,
            abkant_long_single_bend_qty INTEGER NOT NULL DEFAULT 0,
            abkant_long_double_bend_qty INTEGER NOT NULL DEFAULT 0,
            abkant_short_single_bend_qty INTEGER NOT NULL DEFAULT 0,
            abkant_short_double_bend_qty INTEGER NOT NULL DEFAULT 0,
            abkant_manual_override INTEGER NOT NULL DEFAULT 0,
            bend_type TEXT NOT NULL DEFAULT '',
            calculated_area_mm2 REAL NOT NULL DEFAULT 0,
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
        CREATE TABLE IF NOT EXISTS other_work_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tarih TEXT NOT NULL,
            operator_ismi TEXT NOT NULL,
            calisma_tipi TEXT NOT NULL,
            calisma_saati REAL NOT NULL,
            is_aciklamasi TEXT NOT NULL,
            participants_text TEXT NOT NULL DEFAULT '',
            neden TEXT,
            notlar TEXT,
            created_at TEXT NOT NULL
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS app_settings (
            setting_key TEXT PRIMARY KEY,
            setting_value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    cur.execute(
        """
        INSERT OR IGNORE INTO app_settings (
            setting_key, setting_value, updated_at
        ) VALUES ('daily_m2_target', '500', ?)
        """,
        (datetime.now().isoformat(timespec="seconds"),),
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

    cur.execute("PRAGMA table_info(operation_work_logs)")
    operation_log_columns = {row[1] for row in cur.fetchall()}
    operation_log_migrations = {
        "operation_hours": "REAL NOT NULL DEFAULT 0",
        "laser_plate_qty": "INTEGER NOT NULL DEFAULT 0",
        "material_type": "TEXT NOT NULL DEFAULT ''",
        "thickness_mm": "REAL NOT NULL DEFAULT 0",
        "laser_long_edge_qty": "INTEGER NOT NULL DEFAULT 0",
        "laser_short_edge_qty": "INTEGER NOT NULL DEFAULT 0",
        "laser_long_per_team": "INTEGER NOT NULL DEFAULT 2",
        "laser_short_per_team": "INTEGER NOT NULL DEFAULT 2",
        "abkant_work_mode": "TEXT NOT NULL DEFAULT ''",
        "abkant_coworker": "TEXT NOT NULL DEFAULT ''",
        "abkant_team_qty": "INTEGER NOT NULL DEFAULT 0",
        "abkant_teams_per_piece": "INTEGER NOT NULL DEFAULT 1",
        "abkant_long_bend_qty": "INTEGER NOT NULL DEFAULT 0",
        "abkant_short_bend_qty": "INTEGER NOT NULL DEFAULT 0",
        "abkant_long_single_bend_qty": "INTEGER NOT NULL DEFAULT 0",
        "abkant_long_double_bend_qty": "INTEGER NOT NULL DEFAULT 0",
        "abkant_short_single_bend_qty": "INTEGER NOT NULL DEFAULT 0",
        "abkant_short_double_bend_qty": "INTEGER NOT NULL DEFAULT 0",
        "abkant_manual_override": "INTEGER NOT NULL DEFAULT 0",
        "bend_type": "TEXT NOT NULL DEFAULT ''",
        "calculated_area_mm2": "REAL NOT NULL DEFAULT 0",
        "participants_text": "TEXT NOT NULL DEFAULT ''",
        "piece_weight_kg": "REAL NOT NULL DEFAULT 0",
    }
    for column_name, column_definition in operation_log_migrations.items():
        if column_name not in operation_log_columns:
            cur.execute(
                f"ALTER TABLE operation_work_logs "
                f"ADD COLUMN {column_name} {column_definition}"
            )

    cur.execute("PRAGMA table_info(other_work_logs)")
    other_work_columns = {row[1] for row in cur.fetchall()}
    if "participants_text" not in other_work_columns:
        cur.execute(
            "ALTER TABLE other_work_logs "
            "ADD COLUMN participants_text TEXT NOT NULL DEFAULT ''"
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
    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_other_work_date "
        "ON other_work_logs(tarih, operator_ismi)"
    )

    conn.commit()
    conn.close()


def insert_session(session_data, operation_rows):
    conn = get_db_connection()
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
    _clear_data_caches()
    conn.close()
    return session_id




def production_status(siparis_adedi: int, tamamlanan_adet: int) -> str:
    """Sipariş ve tamamlanan adede göre okunabilir üretim durumu döndürür."""
    siparis = max(int(siparis_adedi or 0), 0)
    tamamlanan = max(int(tamamlanan_adet or 0), 0)
    if tamamlanan <= 0:
        return "Başlanmadı"
    if tamamlanan > siparis and siparis > 0:
        return "Fazla üretim"
    if tamamlanan == siparis and siparis > 0:
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
    conn = get_db_connection()
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
        _clear_data_caches()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@st.cache_data(ttl=300, show_spinner=False)
def get_update_history(session_id: int) -> pd.DataFrame:
    conn = get_db_connection()
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
    conn = get_db_connection()
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
        _clear_data_caches()
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

    # Ayar dosyasındaki mevcut kombinasyonların sonuna Plazma'yı ekler.
    # Mevcut ayar dosyasında 14 kombinasyon bulunduğu için bu kayıt
    # uygulamada 15. kombinasyon olarak görünür. Aynı kayıt Excel'e daha
    # sonra eklenirse ikinci kez oluşturulmaz.
    plasma_exists = any(
        any(_normalize_header(operation) == "plazma" for operation in combo.get("operasyonlar", []))
        for combo in combinations
    )
    if not plasma_exists:
        combinations.append(
            {
                "ad": "Plazma",
                "operasyonlar": ["plazma"],
            }
        )

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
    conn = get_db_connection()
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
        _clear_data_caches()
        return len(rows)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@st.cache_data(ttl=300, show_spinner=False)
def get_production_output_summary(oc_no: str | None = None) -> pd.DataFrame:
    conn = get_db_connection()
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
    result["overproduction_qty"] = (result["produced_qty"] - result["requested_qty"]).clip(lower=0)
    result["requested_area_mm2"] = result["unit_area_mm2"] * result["requested_qty"]
    result["produced_area_mm2"] = result["unit_area_mm2"] * result["produced_qty"]
    result["remaining_area_mm2"] = result["unit_area_mm2"] * result["remaining_qty"]
    result["overproduction_area_mm2"] = result["unit_area_mm2"] * result["overproduction_qty"]
    result["_pos_no"] = result["pos"].str.extract(r"(\d+)")[0].fillna(0).astype(int)
    return result.sort_values(["oc_no", "_pos_no"]).drop(columns=["_pos_no"]).reset_index(drop=True)



def delete_production_output_oc(oc_no: str) -> dict:
    """Bir OC'ye ait yüklenen çıktı satırlarını ve günlük çıktı loglarını siler."""
    conn = get_db_connection()
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
        _clear_data_caches()

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
        f"{_format_m2(selected_summary['istenen_alan_mm2'])} m²"
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

    conn = get_db_connection()
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
        _clear_data_caches()
        return len(selected)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@st.cache_data(ttl=300, show_spinner=False)
def get_production_output_history(oc_no: str) -> pd.DataFrame:
    conn = get_db_connection()
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
                ROUND(l.produced_area_mm2 / 1000000.0, 1) AS "Üretilen Alan (m²)",
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
    """mm² olarak saklanan alanı m² ve tek ondalıkla gösterir."""
    numeric = float(value_mm2 or 0) / 1_000_000
    return f"{numeric:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _to_m2(value_mm2) -> float:
    return round(float(value_mm2 or 0) / 1_000_000, 1)


def _operation_kind(operation_name: str) -> str:
    normalized = _normalize_header(operation_name)
    # Cut Out, normal Laser'dan önce kontrol edilir. Böylece "Laser Cut Out",
    # "Laser Cut-Out" ve "Lazer Cutout" gibi yazımlar yalnızca adet olur.
    laser_word = ("laser" in normalized or "lazer" in normalized)
    cutout_word = (
        "cutout" in normalized
        or "cutoff" in normalized
        or ("cut" in normalized and "out" in normalized)
    )
    if laser_word and cutout_word:
        return "laser_cut_out"
    if laser_word:
        return "laser"
    if "abkant" in normalized or "abkand" in normalized:
        return "abkant"
    if "kaynak" in normalized:
        return "kaynak"
    if "cita" in normalized:
        return "cita"
    if "boya" in normalized:
        return "boya"
    if "paket" in normalized:
        return "paketleme"
    if "sevkiyat" in normalized:
        return "sevkiyat"
    return "other"


def _entries_total_hours(entries: list[dict]) -> float:
    return round(
        sum(max(float(entry.get("operation_hours", 0) or 0), 0.0) for entry in entries),
        2,
    )


def _split_participants(value) -> list[str]:
    """Metin veya liste olarak gelen çalışan adlarını tekilleştirir."""
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        raw_values = list(value)
    else:
        try:
            if pd.isna(value):
                return []
        except (TypeError, ValueError):
            pass
        raw_values = re.split(r"[,;|]", str(value))

    result = []
    for part in raw_values:
        name = str(part).strip()
        if name and name not in result:
            result.append(name)
    return result


def _join_participants(values) -> str:
    result = []
    for value in values or []:
        name = str(value).strip()
        if name and name not in result:
            result.append(name)
    return ", ".join(result)


@st.cache_data(ttl=300, show_spinner=False)
def get_daily_m2_target() -> float:
    conn = get_db_connection()
    try:
        row = conn.execute(
            "SELECT setting_value FROM app_settings WHERE setting_key = ?",
            ("daily_m2_target",),
        ).fetchone()
        return max(float(row[0] if row else 500), 0.0)
    except (TypeError, ValueError):
        return 500.0
    finally:
        conn.close()


def save_daily_m2_target(target_m2: float):
    target = max(float(target_m2 or 0), 0.0)
    conn = get_db_connection()
    try:
        conn.execute(
            """
            INSERT INTO app_settings (setting_key, setting_value, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(setting_key) DO UPDATE SET
                setting_value = excluded.setting_value,
                updated_at = excluded.updated_at
            """,
            (
                "daily_m2_target",
                str(target),
                datetime.now().isoformat(timespec="seconds"),
            ),
        )
        conn.commit()
        _clear_data_caches()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def save_other_work_log(data: dict) -> int:
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO other_work_logs (
                tarih, operator_ismi, calisma_tipi, calisma_saati,
                is_aciklamasi, participants_text, neden, notlar, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                str(data["tarih"]),
                str(data["operator_ismi"]).strip(),
                str(data["calisma_tipi"]),
                float(data["calisma_saati"]),
                str(data["is_aciklamasi"]).strip(),
                str(data.get("participants_text", "") or ""),
                str(data.get("neden", "") or ""),
                str(data.get("notlar", "") or ""),
                datetime.now().isoformat(timespec="seconds"),
            ),
        )
        record_id = int(cur.lastrowid)

        parsed_date = pd.to_datetime(data["tarih"], errors="coerce")
        weekend_work = bool(
            not pd.isna(parsed_date) and parsed_date.weekday() >= 5
        )
        operation_hours = float(cur.execute(
            """
            SELECT COALESCE(SUM(calisma_saati), 0)
            FROM operation_batches
            WHERE operator_ismi = ? AND tarih = ?
            """,
            (str(data["operator_ismi"]).strip(), str(data["tarih"])),
        ).fetchone()[0] or 0)
        other_hours = float(cur.execute(
            """
            SELECT COALESCE(SUM(calisma_saati), 0)
            FROM other_work_logs
            WHERE operator_ismi = ? AND tarih = ?
            """,
            (str(data["operator_ismi"]).strip(), str(data["tarih"])),
        ).fetchone()[0] or 0)
        if weekend_work or operation_hours + other_hours > 9.0:
            cur.execute(
                """
                UPDATE operation_batches
                SET calisma_tipi = 'Mesaili'
                WHERE operator_ismi = ? AND tarih = ?
                """,
                (str(data["operator_ismi"]).strip(), str(data["tarih"])),
            )
            cur.execute(
                """
                UPDATE other_work_logs
                SET calisma_tipi = 'Mesaili'
                WHERE operator_ismi = ? AND tarih = ?
                """,
                (str(data["operator_ismi"]).strip(), str(data["tarih"])),
            )

        conn.commit()
        _clear_data_caches()
        return record_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@st.cache_data(ttl=300, show_spinner=False)
def get_other_work_logs() -> pd.DataFrame:
    conn = get_db_connection()
    try:
        return pd.read_sql_query(
            """
            SELECT
                id, tarih, operator_ismi, calisma_tipi, calisma_saati,
                is_aciklamasi, participants_text, neden, notlar, created_at
            FROM other_work_logs
            ORDER BY tarih DESC, id DESC
            """,
            conn,
        )
    finally:
        conn.close()


def expand_other_work_attribution(other_logs: pd.DataFrame) -> pd.DataFrame:
    """Diğer çalışma ekibindeki kişileri çalışan bazlı saat özetlerine dahil eder."""
    if other_logs is None or other_logs.empty:
        return (
            other_logs.copy()
            if isinstance(other_logs, pd.DataFrame)
            else pd.DataFrame()
        )

    base = other_logs.copy()
    base["ana_operator"] = base.get("operator_ismi", "")
    base["katilim_rolu"] = "Ana çalışan"
    shared_rows = []

    for _, row in base.iterrows():
        main_worker = str(row.get("operator_ismi", "") or "").strip()
        for coworker in _split_participants(row.get("participants_text", "")):
            if not coworker or coworker == main_worker:
                continue
            shared = row.copy()
            shared["ana_operator"] = main_worker
            shared["operator_ismi"] = coworker
            shared["katilim_rolu"] = "Çalışma ekibi"
            shared_rows.append(shared)

    if not shared_rows:
        return base
    return pd.concat([base, pd.DataFrame(shared_rows)], ignore_index=True, sort=False)


def delete_other_work_log(record_id: int):
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM other_work_logs WHERE id = ?", (int(record_id),))
        conn.commit()
        _clear_data_caches()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_registered_daily_hours(operator_name: str, production_date) -> float:
    if not operator_name or production_date is None:
        return 0.0
    conn = get_db_connection()
    try:
        row = conn.execute(
            """
            SELECT COALESCE(SUM(calisma_saati), 0)
            FROM (
                SELECT calisma_saati
                FROM operation_batches
                WHERE operator_ismi = ? AND tarih = ?
                UNION ALL
                SELECT calisma_saati
                FROM other_work_logs
                WHERE operator_ismi = ? AND tarih = ?
            )
            """,
            (
                str(operator_name), str(production_date),
                str(operator_name), str(production_date),
            ),
        ).fetchone()
        return float(row[0] or 0)
    finally:
        conn.close()


def determine_work_type(
    production_date,
    declared_type: str,
    current_hours: float,
    existing_hours: float = 0.0,
) -> str:
    parsed_date = pd.to_datetime(production_date, errors="coerce")
    is_weekend = bool(not pd.isna(parsed_date) and parsed_date.weekday() >= 5)
    cumulative_hours = float(existing_hours or 0) + float(current_hours or 0)
    if is_weekend or cumulative_hours > 9.0 or str(declared_type) == "Mesaili":
        return "Mesaili"
    if str(declared_type) == "Yarı zamanlı":
        return "Yarı zamanlı"
    return "Tam zamanlı"


def render_pos_area_overview(
    boy_mm: float,
    en_mm: float,
    unit_area_mm2: float,
    requested_qty: int,
    completed_qty: int,
):
    requested_area_mm2 = float(unit_area_mm2) * int(requested_qty)
    remaining_qty = max(int(requested_qty) - int(completed_qty), 0)
    overproduction_qty = max(int(completed_qty) - int(requested_qty), 0)
    actual_completion_pct = (
        max(int(completed_qty) / int(requested_qty) * 100, 0)
        if int(requested_qty) > 0
        else 0
    )
    ring_completion_pct = min(actual_completion_pct, 100)
    ring_degrees = ring_completion_pct * 3.6
    completion_detail = (
        f"{overproduction_qty} adet fazla üretildi"
        if overproduction_qty > 0
        else f"{remaining_qty} adet kaldı"
    )
    balance_label = "Fazla üretim alanı" if overproduction_qty > 0 else "Kalan alan"
    balance_area_mm2 = unit_area_mm2 * (
        overproduction_qty if overproduction_qty > 0 else remaining_qty
    )
    balance_detail = (
        "Sipariş üzeri sağlam üretim"
        if overproduction_qty > 0
        else "Kalan üretim ihtiyacı"
    )

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
                    Birim alan · {_format_m2(unit_area_mm2)} m²
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
                        %{actual_completion_pct:.0f}
                    </div>
                    <div>
                        <div class="area-stat-label">Tamamlanma</div>
                        <div class="area-stat-value">
                            {int(completed_qty)} / {int(requested_qty)}
                        </div>
                        <div class="area-stat-detail">
                            {completion_detail}
                        </div>
                    </div>
                </div>
                <div class="area-stat-card">
                    <div class="area-stat-icon">⌁</div>
                    <div class="area-stat-label">{balance_label}</div>
                    <div class="area-stat-value">
                        {_format_m2(balance_area_mm2)} m²
                    </div>
                    <div class="area-stat-detail">
                        {balance_detail}
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
                    Net üretim alanı
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




def reset_new_operation_wizard_state():
    """Yeni Kayıt sayfasına girildiğinde önceki yarım formu güvenle temizler."""
    current_version = int(st.session_state.get("operation_form_version", 0))
    st.session_state["operation_form_version"] = current_version + 1

    removable_prefixes = (
        "operation_wizard_step_",
        "operation_wizard_data_",
        "op_worker_",
        "op_date_",
        "op_work_type_",
        "op_reason_",
        "op_oc_",
        "op_positions_",
        "change_plan_",
        "op_combo_",
        "same_operations_",
        "common_ops_",
        "selected_ops_",
        "processed_",
        "opfire_",
        "ophours_",
        "laserplates_",
        "material_",
        "thickness_",
        "laser_lot_count_",
        "abkantmode_",
        "abkantcoworker_",
        "abkantteamqty_",
        "abkantteamsperpiece_",
        "abkantlongbendqty_",
        "abkantshortbendqty_",
        "abkantlongsinglebendqty_",
        "abkantlongdoublebendqty_",
        "abkantshortsinglebendqty_",
        "abkantshortdoublebendqty_",
        "abkantmanualoverride_",
        "participants_",
        "pieceweight_",
        "bendtype_",
        "show_other_operations_",
        "add_laser_lot_",
        "remove_laser_lot_",
        "fill_",
        "fill_cutout_",
    )
    for key in list(st.session_state.keys()):
        if any(str(key).startswith(prefix) for prefix in removable_prefixes):
            st.session_state.pop(key, None)


def handle_main_page_change():
    """Sayfa geçişini izler; Yeni Kayıt'a her girişte temiz bir form açar."""
    selected_page = str(st.session_state.get("main_page_selector", "Yeni Kayıt"))
    previous_page = st.session_state.get("_previous_main_page")
    if selected_page != previous_page:
        st.session_state["_previous_main_page"] = selected_page
        if selected_page == "Yeni Kayıt":
            reset_new_operation_wizard_state()


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

    conn = get_db_connection()
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
        _clear_data_caches()
    finally:
        conn.close()


@st.cache_data(ttl=300, show_spinner=False)
def get_performance_targets() -> pd.DataFrame:
    conn = get_db_connection()
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
    conn = get_db_connection()
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
        _clear_data_caches()
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


@st.cache_data(ttl=300, show_spinner=False)
def get_historical_efficiency_reference() -> dict:
    conn = get_db_connection()
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
    timed_good_area = sum(
        max(int(entry.get("good_qty", 0)), 0)
        * float(entry.get("unit_area_mm2", 0))
        for entry in entries
        if float(entry.get("operation_hours", 0) or 0) > 0
    )

    safe_hours = float(work_hours) if float(work_hours) > 0 else 0.0
    qty_per_hour = total_good / safe_hours if safe_hours else 0.0
    area_per_hour = timed_good_area / safe_hours if safe_hours else 0.0
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
        "Her işlem için adet/saat, m²/saat ve fire sınırı belirle. "
        "Hedefi 0 bırakırsan gidişat geçmiş kayıt ortalamasına göre hesaplanır."
    )

    operations = all_operations_from_combinations(combinations)

    conn = get_db_connection()
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

    display_targets = targets.copy()
    display_targets["target_area_per_hour"] = (
        display_targets["target_area_per_hour"] / 1_000_000
    ).round(1)

    editor = display_targets[
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
            "target_area_per_hour": "Hedef m²/Saat",
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
            "Hedef m²/Saat": st.column_config.NumberColumn(
                min_value=0.0,
                step=0.1,
                format="%.1f",
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
                "Hedef m²/Saat": "target_area_per_hour",
                "Fire Sınırı %": "fire_limit_pct",
                "Yavaş Altı %": "slow_limit_pct",
                "Hızlı Üstü %": "fast_limit_pct",
            }
        )
        save_frame["target_area_per_hour"] = (
            pd.to_numeric(
                save_frame["target_area_per_hour"], errors="coerce"
            ).fillna(0)
            * 1_000_000
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



@st.cache_data(ttl=300, show_spinner=False)
def get_worker_competencies(worker_name: str) -> list[str]:
    if not worker_name:
        return []
    conn = get_db_connection()
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
    conn = get_db_connection()
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
        _clear_data_caches()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@st.cache_data(ttl=300, show_spinner=False)
def get_competency_table() -> pd.DataFrame:
    conn = get_db_connection()
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


@st.cache_data(ttl=300, show_spinner=False)
def get_item_operation_plan(item_id: int) -> pd.DataFrame:
    conn = get_db_connection()
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


@st.cache_data(ttl=300, show_spinner=False)
def get_operation_progress(item_ids: list[int] | None = None) -> pd.DataFrame:
    conn = get_db_connection()
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
    result["completed_qty"] = result["operation_good_qty"].clip(lower=0).astype(int)
    result["remaining_qty"] = (
        result["requested_qty"] - result["completed_qty"]
    ).clip(lower=0).astype(int)
    result["overproduction_qty"] = (
        result["completed_qty"] - result["requested_qty"]
    ).clip(lower=0).astype(int)
    result["completion_pct"] = (
        result["completed_qty"]
        / result["requested_qty"].replace(0, 1)
        * 100
    ).round(1)
    return result


@st.cache_data(ttl=300, show_spinner=False)
def get_operation_history(oc_no: str | None = None, pos: str | None = None) -> pd.DataFrame:
    conn = get_db_connection()
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
                COALESCE(NULLIF(i.project_name, ''), 'OC ' || i.oc_no) AS proje_adi,
                i.pos AS POS,
                i.boy_mm,
                i.en_mm,
                w.operation_name AS operasyon,
                w.operation_hours AS operasyon_saati,
                w.laser_plate_qty AS laser_plaka_adedi,
                w.material_type AS malzeme,
                w.thickness_mm AS kalinlik_mm,
                w.laser_long_edge_qty AS laser_uzun_kenar,
                w.laser_short_edge_qty AS laser_kisa_kenar,
                w.laser_long_per_team AS takim_uzun_kenar,
                w.laser_short_per_team AS takim_kisa_kenar,
                w.abkant_work_mode AS abkant_calisma_sekli,
                w.abkant_coworker AS beraber_calistigi,
                w.abkant_team_qty AS abkant_takim_sayisi,
                w.abkant_teams_per_piece AS adet_basina_takim,
                w.abkant_long_bend_qty AS abkant_uzun_kenar_bukum,
                w.abkant_short_bend_qty AS abkant_kisa_kenar_bukum,
                w.abkant_long_single_bend_qty AS abkant_uzun_tek_bukum,
                w.abkant_long_double_bend_qty AS abkant_uzun_cift_bukum,
                w.abkant_short_single_bend_qty AS abkant_kisa_tek_bukum,
                w.abkant_short_double_bend_qty AS abkant_kisa_cift_bukum,
                w.abkant_manual_override AS abkant_manuel_adet,
                w.participants_text AS beraber_calisanlar,
                w.piece_weight_kg AS parca_agirligi_kg,
                w.bend_type AS bukum_turu,
                w.calculated_area_mm2 AS hesaplanan_alan_mm2,
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
        result = pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()

    if result.empty:
        return result

    # Bütün üretim etaplarında alan, sağlam adet × POS birim alanı olarak hesaplanır.
    # Laser Cut Out ekranında yalnızca adet girilir; m² değeri otomatik oluşur.
    result["saglam_alan_m2"] = (
        pd.to_numeric(result["saglam_alan_mm2"], errors="coerce").fillna(0)
        / 1_000_000
    ).round(1)
    result["islenen_alan_m2"] = (
        pd.to_numeric(result["islenen_alan_mm2"], errors="coerce").fillna(0)
        / 1_000_000
    ).round(1)
    return result


def expand_worker_attribution(history: pd.DataFrame) -> pd.DataFrame:
    """
    Abkant ile ekip seçilebilen tüm operasyonlardaki arkadaşları kişi bazlı analizlere dahil eder.

    Şirket toplamı iki kez sayılmamalıdır. Bu genişletilmiş veri yalnızca
    çalışan/kişi bazlı özetlerde kullanılır; ekipteki her kişi aynı işin katkısını
    kendi satırında görür.
    """
    if history is None or history.empty:
        return history.copy() if isinstance(history, pd.DataFrame) else pd.DataFrame()

    base = history.copy()
    if "ana_operator" not in base.columns:
        base["ana_operator"] = base.get("operator_ismi", "")
    base["katilim_rolu"] = "Ana çalışan"

    shared_rows = []
    for _, row in base.iterrows():
        main_worker = str(row.get("operator_ismi", "") or "").strip()
        kind = _operation_kind(str(row.get("operasyon", "") or ""))
        coworkers = []

        if kind == "abkant":
            mode = str(row.get("abkant_calisma_sekli", "") or "").strip()
            if mode == "Biriyle beraber çalıştı":
                coworkers.extend(
                    _split_participants(row.get("beraber_calistigi", ""))
                )
        coworkers.extend(
            _split_participants(row.get("beraber_calisanlar", ""))
        )

        unique_coworkers = []
        for coworker in coworkers:
            if coworker and coworker != main_worker and coworker not in unique_coworkers:
                unique_coworkers.append(coworker)

        for coworker in unique_coworkers:
            shared = row.copy()
            shared["ana_operator"] = main_worker
            shared["operator_ismi"] = coworker
            shared["katilim_rolu"] = "Beraber çalışan"
            shared_rows.append(shared)

    if not shared_rows:
        return base
    return pd.concat([base, pd.DataFrame(shared_rows)], ignore_index=True, sort=False)


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

    conn = get_db_connection()
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

        parsed_batch_date = pd.to_datetime(batch_data["tarih"], errors="coerce")
        weekend_work = bool(
            not pd.isna(parsed_batch_date) and parsed_batch_date.weekday() >= 5
        )
        operation_hours = float(cur.execute(
            """
            SELECT COALESCE(SUM(calisma_saati), 0)
            FROM operation_batches
            WHERE operator_ismi = ? AND tarih = ?
            """,
            (batch_data["operator_ismi"], batch_data["tarih"]),
        ).fetchone()[0] or 0)
        other_hours = float(cur.execute(
            """
            SELECT COALESCE(SUM(calisma_saati), 0)
            FROM other_work_logs
            WHERE operator_ismi = ? AND tarih = ?
            """,
            (batch_data["operator_ismi"], batch_data["tarih"]),
        ).fetchone()[0] or 0)
        cumulative_day_hours = operation_hours + other_hours
        if weekend_work or cumulative_day_hours > 9.0:
            cur.execute(
                """
                UPDATE operation_batches
                SET calisma_tipi = 'Mesaili'
                WHERE operator_ismi = ? AND tarih = ?
                """,
                (batch_data["operator_ismi"], batch_data["tarih"]),
            )
            cur.execute(
                """
                UPDATE other_work_logs
                SET calisma_tipi = 'Mesaili'
                WHERE operator_ismi = ? AND tarih = ?
                """,
                (batch_data["operator_ismi"], batch_data["tarih"]),
            )

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
            combination_name = str(plan.get("combination_name", "") or "").strip()

            # Aynı operasyon adı plana birden fazla kez geldiyse SQLite'taki
            # UNIQUE(item_id, operation_name) kuralı kayıt işlemini durduruyordu.
            # Operasyon sırasını koruyarak tekrarları burada temizliyoruz.
            unique_operations = []
            seen_operation_names = set()
            for raw_operation in plan.get("operations", []):
                operation = " ".join(str(raw_operation or "").split())
                if not operation:
                    continue
                operation_key = operation.casefold()
                if operation_key in seen_operation_names:
                    continue
                seen_operation_names.add(operation_key)
                unique_operations.append(operation)

            cur.execute(
                "UPDATE production_output_items SET combination_name = ? WHERE id = ?",
                (combination_name, int(item_id)),
            )
            cur.execute(
                "DELETE FROM pos_operation_plan WHERE item_id = ?",
                (int(item_id),),
            )
            cur.executemany(
                """
                INSERT OR REPLACE INTO pos_operation_plan (
                    item_id, combination_name, operation_order,
                    operation_name, created_at
                ) VALUES (?, ?, ?, ?, ?)
                """,
                [
                    (
                        int(item_id),
                        combination_name,
                        index,
                        operation,
                        created_at,
                    )
                    for index, operation in enumerate(unique_operations, start=1)
                ],
            )

        for entry in entries:
            item_id = int(entry["item_id"])
            operation_name = str(entry["operation_name"])
            processed_qty = int(entry["processed_qty"])
            fire_qty = int(entry["fire_qty"])
            operation_kind = _operation_kind(operation_name)

            if operation_kind == "abkant":
                # Önceki günlerde girilmiş uzun ve kısa kenarları ayrı ayrı toplar.
                # Bir ürün, hem uzun hem kısa kenarı tamamlandığında abkanttan geçmiş sayılır.
                cur.execute(
                    """
                    SELECT
                        COALESCE(SUM(abkant_long_single_bend_qty + abkant_long_double_bend_qty), 0),
                        COALESCE(SUM(abkant_short_single_bend_qty + abkant_short_double_bend_qty), 0)
                    FROM operation_work_logs
                    WHERE item_id = ? AND operation_name = ?
                    """,
                    (item_id, operation_name),
                )
                previous_long, previous_short = cur.fetchone()
                new_long = (
                    int(entry.get("abkant_long_single_bend_qty", 0) or 0)
                    + int(entry.get("abkant_long_double_bend_qty", 0) or 0)
                )
                new_short = (
                    int(entry.get("abkant_short_single_bend_qty", 0) or 0)
                    + int(entry.get("abkant_short_double_bend_qty", 0) or 0)
                )
                cur.execute(
                    "SELECT requested_qty FROM production_output_items WHERE id = ?",
                    (item_id,),
                )
                requested_for_item = int(cur.fetchone()[0] or 0)
                completed_before = min(
                    int(previous_long or 0), int(previous_short or 0), requested_for_item
                )
                completed_after = min(
                    int(previous_long or 0) + new_long,
                    int(previous_short or 0) + new_short,
                    requested_for_item,
                )
                processed_qty = max(completed_after - completed_before, 0)

                has_partial_work = any([
                    new_long > 0,
                    new_short > 0,
                    float(entry.get("operation_hours", 0) or 0) > 0,
                    fire_qty > 0,
                ])
                if not has_partial_work:
                    continue
            elif processed_qty <= 0:
                continue

            if fire_qty < 0 or fire_qty > processed_qty:
                if operation_kind == "abkant" and processed_qty == 0 and fire_qty > 0:
                    raise ValueError(
                        f"{operation_name}: Bu kenar girişleriyle henüz tamamlanan ürün oluşmadığı "
                        "için fire adedi girilemez."
                    )
                raise ValueError(f"{operation_name}: Fire, tamamlanan adetten büyük olamaz.")

            good_qty = processed_qty - fire_qty

            cur.execute(
                """
                INSERT INTO operation_work_logs (
                    batch_id, item_id, operation_name, processed_qty,
                    fire_qty, good_qty, operation_hours, laser_plate_qty,
                    material_type, thickness_mm, laser_long_edge_qty,
                    laser_short_edge_qty, laser_long_per_team,
                    laser_short_per_team, abkant_work_mode,
                    abkant_coworker, abkant_team_qty,
                    abkant_teams_per_piece, abkant_long_bend_qty,
                    abkant_short_bend_qty, abkant_long_single_bend_qty,
                    abkant_long_double_bend_qty, abkant_short_single_bend_qty,
                    abkant_short_double_bend_qty, abkant_manual_override,
                    participants_text, piece_weight_kg, bend_type,
                    calculated_area_mm2, fire_note, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    batch_id,
                    item_id,
                    operation_name,
                    processed_qty,
                    fire_qty,
                    good_qty,
                    float(entry.get("operation_hours", 0) or 0),
                    int(entry.get("laser_plate_qty", 0) or 0),
                    str(entry.get("material_type", "") or ""),
                    float(entry.get("thickness_mm", 0) or 0),
                    int(entry.get("laser_long_edge_qty", 0) or 0),
                    int(entry.get("laser_short_edge_qty", 0) or 0),
                    int(entry.get("laser_long_per_team", 2) or 2),
                    int(entry.get("laser_short_per_team", 2) or 2),
                    str(entry.get("abkant_work_mode", "") or ""),
                    str(entry.get("abkant_coworker", "") or ""),
                    int(entry.get("abkant_team_qty", 0) or 0),
                    max(int(entry.get("abkant_teams_per_piece", 1) or 1), 1),
                    int(entry.get("abkant_long_bend_qty", 0) or 0),
                    int(entry.get("abkant_short_bend_qty", 0) or 0),
                    int(entry.get("abkant_long_single_bend_qty", 0) or 0),
                    int(entry.get("abkant_long_double_bend_qty", 0) or 0),
                    int(entry.get("abkant_short_single_bend_qty", 0) or 0),
                    int(entry.get("abkant_short_double_bend_qty", 0) or 0),
                    int(bool(entry.get("abkant_manual_override", False))),
                    str(entry.get("participants_text", "") or ""),
                    float(entry.get("piece_weight_kg", 0) or 0),
                    str(entry.get("bend_type", "") or ""),
                    (
                        float(good_qty) * float(entry.get("unit_area_mm2", 0) or 0)
                        if operation_kind == "abkant"
                        else float(entry.get("calculated_area_mm2", float(good_qty) * float(entry.get("unit_area_mm2", 0) or 0)) or 0)
                    ),
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
                stage_totals.append(operation_good)

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
        _clear_data_caches()
        return batch_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@st.cache_data(ttl=300, show_spinner=False)
def get_overtime_summary(target_date: str | None = None) -> pd.DataFrame:
    """Eski ve yeni kayıt sistemindeki mesaileri tek özet içinde gösterir."""
    conn = get_db_connection()
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

        other_query = """
            SELECT
                tarih,
                operator_ismi,
                calisma_saati AS toplam_saat,
                'Diğer Çalışma' AS OC,
                '-' AS POS,
                0 AS uretilen_adet,
                0 AS fire_adedi
            FROM other_work_logs
            WHERE calisma_tipi = 'Mesaili'
        """
        other_params = []
        if target_date:
            other_query += " AND tarih = ?"
            other_params.append(str(target_date))
        other = pd.read_sql_query(other_query, conn, params=other_params)
    finally:
        conn.close()

    combined = pd.concat([legacy, modern, other], ignore_index=True)
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
            width: 48px;
            height: 48px;
            min-width: 48px;
            overflow: hidden;
            display: grid;
            place-items: center;
            border-radius: 12px;
            background: #06A5DA;
            border: 1px solid rgba(255,255,255,.22);
            box-shadow: 0 11px 22px rgba(21, 138, 181, 0.28);
        }

        .sidebar-mark img {
            width: 100%;
            height: 100%;
            display: block;
            object-fit: cover;
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
            width: 86px;
            height: 86px;
            min-width: 86px;
            overflow: hidden;
            display: grid;
            place-items: center;
            border-radius: 18px;
            background: #06A5DA;
            border: 1px solid rgba(255,255,255,.20);
            box-shadow: 0 14px 28px rgba(0,0,0,.18);
        }

        .durlum-logo-box img {
            width: 100%;
            height: 100%;
            display: block;
            object-fit: cover;
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

        /* Sayfa değişirken önceki ekranın soluk biçimde altta kalmasını engelle. */
        [data-testid="stMain"] [data-stale="true"],
        section.main [data-stale="true"] {
            opacity: 0 !important;
            pointer-events: none !important;
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
        f'<div class="durlum-logo-box"><img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wgARCADIAMgDASIAAhEBAxEB/8QAHAABAAMBAQEBAQAAAAAAAAAAAAYHCAUEAwIB/8QAGwEBAAIDAQEAAAAAAAAAAAAAAAQFAgMGAQf/2gAMAwEAAhADEAAAAfCPqPw8TXTIhTSlHQbSOC0owAAAAAB78cvAl/h0So8JMMBYNfWDAtL2znozOfOdjFB2XzoAAAAAB3OH99e2VuT54VjyxY1KZQ3S1TfR7sduPct3EwznozOc2useNXDBteyDzviW1nhmC4Ks0lKgQDyzn21HQ5f5dg192/zQJMIAAABqXLWpea7Ll1zY1cw7G7M56Mzn5loODTmDQLTj21UttZ6s26SzbpKbXU5YtdVtMr5tX31+V7ywSIoAAADUuWtS812XLrmxq5h2N2Zz0ZnPzLQcGnMGgWnHtqpba2as26SzbpKZXU7VVq1VfcsFnSgAAAANO5ivGh6mVwC2nM9qznozOdlTaC4H9lFVewGfGGzNuks26Suecp2qrVqq+5YLOlAAAAAfT5iR/qNI8uS8PzM9fp7kaEmRlhs91lV9qCl6TOMWtSq7SjCdWAAAAAAAAAAdfUGX9Qcj39OVXalZXXN/B6fjZ0v4er4Pfw9P9898r1f08j0ef3EPfAAAAAOvqDL+oOR7+nJhD5hjn26Tuyk/cdCUDf1A6ZN1dTl+WqvO84v2x2dSF9Dpbo2Wh9G+PgAAAAdfUGX9Qcj39OTCHzDHPt0ndlJ+46EoG/qB0ybqi8oi8Sc6vK6u3TFp9AZ957lod58tAAAAA6+oMv6g5Hv6cmEPmGOfbpO7KT9x0JQN/UDpk3VF5RF4k51OX1NumMT2BT3z3LQ7z5aAAAAB19QZb0Dyvc1/LIH4tmm/Y4/dB1cjoG26ZtaO9Yv6o9Fm9fqRPo7NPjntXTJ7nkdt81AAAAAAAAAAAAA//8QAKRAAAQMDAwMEAwEBAAAAAAAABgQFNQACAwEwMxY0NhAUMUATFSARIv/aAAgBAQABBQL7gmjwrXPp1tohT40rxvo03uslmLJraoS2ZbP4CJeime30P/aO/KgyVrfrkcv4CJeime38Oa9Pl1Uo81KFv5Mfowjt7zWgQg/xsHErSoopnsIYgvxEg6laUAuyJ3jTolvrElsvd+iW+uiW+swMkutcm/I1q9sfwaJ2ZcrtQJGIiterqKZ5N2xvEAXxSfyOlTgmRa2X6ZLTeY22yNJYMFkKKZ5N2xvEAXxSfyOjznYYY3mNttjiWDBZCimeTdsbxAF8Un8jo857VeezTJkvy67bbHEsGCyFFM8m7Y3iAL4pP5HR5zbrZGksGCyFFM8m7Y3iAL4pP5HR5zbrLf8AkaHpNesag9tVIltFM8l7YjbMjs3CrMpadKT+R0ec26GOdmZD6lM8OOFi9r9U/kdHnNu2X3Y7tCJyt06jcq6jcqUKMirNgU5Ut/UjlXUjlXUjlSfXNmW+1JqfcTjjv+s0StHnP9ZolaPOf6zRK0ec9uDJfpqny6aW2636+2zVdZdZromy66e1zV7XNXtc1XYMlu60StHnOIwTzEiE7RlNtcZ/BKz4FbfttErR5ziME8xIhO0ZTbXGPTp+oRtDna7InNb+uQsbza9J3ON22iVo85xGCeYkQnaMptrjDSGC4YmggTtHOO22iVo85xGCeYkQnaMptrjDSGC4YmggTtHOO22iVo85xGCeYkQnaMptrjDSGC4YmggTtHOO22iVo85xGCeIkQnaMptrjDSGC4YmggTtHOO22iVo85w6/S5kvs0yWIh1E3qaMptrjDSFC4YnggTtHOO223Jbiceom2jFenX5hp/0aL7SVtv06iba6ibaKFWJY7N7834kBU7o1rUKuyNE1v70hUtAe5pUKZe/N+RD9r//xAA3EQABAwEDCQUGBwEAAAAAAAABAAIDBAURcQYSITE0NYGxwRMgM1FyFDAyQWHSECJCU4KR0RX/2gAIAQMBAT8BVPCaiVsTfmq6yZaCMSPcDpu9w1pec0J1M4C8dfxszbYsVlHsrfV0PuInBjrynTi68kcPoPno/CxqJlZORL8IQgs+nnZGGgPOpZR7K31dCq6gpY6N72xi+5WVQ001Gx8kYJ08yrFpIJxL2rAbiv8AnUDyWZgVbC2CofE3UD3smvEkwCrd7wYf6so9lb6uhVo7BJgrF2CPjzKyf+GbFS14s+0pnlt993IKqm9pndLddf3smvEkwCrd7wYf6so9lb6uhVo7BJgrF2CPjzKyf+GbFWzt8nDkO/k9M2OpLHfqCloWS1LKknS1ZR7K31dCgI62mu/S4Kmp2UsQhj1BZP8AwzYq2dvk4ch39S9rqP3D/ZT55ZRc9xPFMnliFzHEcV7ZU/uH+yrFoxVMe4vcNPyNytKLsKt8YJN3nr1e8yb8KTFWwCa+S76cgs0otI1hZjvJZjvJEEa+/k34UmKpt9TYfarU3jTYjmsodj4jqhI2GnEj9QCE0Zi7YH8t1/BS9jW0pI0tI7+TfhSYqm31Nh9qtTeNNiOayh2PiOqrN3O9PRQ7p/h0Vm7tbgevfyb8KTFU2+psPtVqbxpsRzWUOx8R1VZu53p6KHdP8Ois3drcD17+T00cUT89wGlNrIoLYfI4/lOi/gE6WilIe5zSRq0hW9PFJSXMcDp81V1EJoHNDxfm+aiqIRZeZni/M8/orPqIW2c1rni+4/PH3/8A/8QAMxEAAAQEAQkGBwEAAAAAAAAAAAECAwQFEXE0EhMgITEzNUHBFDBRgdHwECIyUoKR4bH/2gAIAQIBAT8BDzhMtm4fIQswRFryEl3BmSSqYJ8jOnxjsMuwk2/VbqXcLTlJoQJk/A/hMolUM18m0xnYt5pS6maeYk2/VbqQhYt9cSlKlnSomEU+3EqShdC/gmcQ60aMhVNQ7ZFp+bLMQzhuspWraelO/oQIbh7t/QSbfqt1IQeLRcTPFr8v8E32t2DcIcXBNpI6U/oh2sy0lvw0p39CBDcPdv6CTb9VupCDxaLiZ4tfl/gm+1uwluER756c4bNTJKLkG4pTbKmCLUYk2/VbqQM1wz9eZGHnlPuG4vaYm+1uwluER759x2dn7C/QS02g6oSRBTTa9akkY7Oz9hfoTOJNhaSJJHchBOZ2HSulLd5Ot4iwlx0hEe+YqQqRjKLxGUnxFa6c63iLB7hjd/UQGDf98hJ8T5A0KcdNCdpmM0sl5umvYEZyGfIthkenOt4iwe4Y3f1EBg3/AHyEnxPkIfGJuHMf+XURuNVctOdbxFg9wxu/qIDBv++Qk+J8hD4xNw5xD8uojcaq5ac4bWtackqg4Zx2XJQRay/oSiJQRpSRlW4lLTiIiqk01CHZcKLSZpPaFsuduysk6ZXURbLhxhmST29//8QAQRAAAQICBgYFCwIEBwAAAAAAAgABAwQQERJyc7ETITBxssExNFGBgxQgIzIzQEFhdIKSIuEFFZGhJEJSYmPR8P/aAAgBAQAGPwL3woceG0QNG71PvZdUhqYhQhYIY1VC25vcHrKxDFrRn2MmeW/hNuC/QcQSJyUUghPLR4OuJAfs7W808J82omu7hb3Cehj69hi3sz6//fJSITDE/oGZzhn6ut+lqlOxysWBgkz2CtD6tltf9PNPCfNqJru4W9wGJDKyY9Dq1ElTA/joYlQv3OzrQwobQIFddltbu/zekjItFAF6rXxfcvXjv9zf9J40Eojk42f1PRNd3CyAneNW7M/rIY0F4lp4jD+p96mdO5+js1WH3rpjfl+yCWevRvH0fzqtVLpjfl+y6Y35fsvRxooF86nRy8TW7fFvi20lBb4ha/rrUWYNqxBq6mUQNFoTDXVXXW1E13cLKFdZBjNk6nfs50Q/q246BaPGCE5dFp0xC7EL62dkOE2b7SVwhyU3d5qYwudE13cLKFdZBjNk6nfs50Q/q246JS6Sk8NkGE2b7SVwhyU3d5qYwudE13cLKFdZBjNk6nfs50Q/q246JS6SqGNEZm+DEqzJzftJ9pK4Q5Kbu81MYXOia7uFlCusgxmydTv2c6If1bcdEpdLltpXCHJTd3mpjC50TXdwsoV1kGM2Tqd+znRD+rbjolLpcttJv/xDkpmDD1mQ6mUYo8A4QvDqZyb50TXdwsoVxlooTtpGJia0pnyhha3ZqsvX20Q/q246JS6XLbeSEXpYXQ3aPmTXdwsoNT+khjYNvMh/Vtx0Sl0uW2YgJxJuh2VXlcRdbNdbNFFilbiF0k6twYhQy7RepdbNdbP+y62f9lDeG/8AiCiNZf8A3Vr2j/mKhfzErRVPY1s/u8ljBnRKXS93ksYM6JS6Xu8ljBnRKXSVYwyJu1mVbwzZrqqFnJ+xl7I/xVRM4v8ANVtCN2ur2R/ivZH+K9if4rXDJt7bWSxgzolLpKDvLNTuCeSg7iyoK4KlMIcvNjRWhsMeGNtjZuzaSWMGdEpdJQd5ZqdwTyUHcWVBXBUphDktPo9L+phs11IY4g4a6nF1FmLGksN6tdSKI0PREL1ONdamsIstpJYwZ0Sl0lB3lmp3BPJQdxZUFcFSmCGS8Rl4jqbutmpq+2SmsIstpJYwZ0Sl0lB3lmp3BPJQdxZUFcFSmCGS8Rl4jqb3Nmpq+2SmsIstpJYwZ0Sl0lB3lmp3BPJQdxZUFcFSmCGS8Rl4jqb3Nmpq+2SmsIstpJYwZ0Sl0lB3lmp3BPJQdxZUFcFSmCGS8Rl4jqb3Nmpm+2SmsIstpJYwZ0Sl0kDf6SJkQE1Yk1TsmjwYbjEbo/U70FcFSmCGS8Rl4jqb3Nmymb7ZKawiy2kqZvZEYou79663DUs8vFGKws9dSOHFZylz16v8rqvyse+tl1uGutw0USCbRAstrZSwFNAxDDFnbuWjgzAxDts9TLRx44wztu9TqZhQpkDiEzVC29TAx4wwnctVamAGaByKGTM3d73/AP/EACkQAAEEAAQFBQEBAQAAAAAAAAEAEVHwECExoSAwQWHxQHGBkdGxweH/2gAIAQEAAT8h9YLRul3/AOi8EUQ2nSBnegbT3tnb7dToB3ITNyOJaXBA+gtRnTt0ZuYZw4L6vxrd36CBrggxAn9IEA4B1j0XHz7rNpSPgLWTk8Zbu/QQ4AxwIlDAX1A+CmsAZv3LrHbtiezVDHLH6QQOZKByVMmQxIMdsN2o75MYY9keUMGyGIKOyKAHQg1zdOy8HTJKImKDwdPB0LI8Wv0w/qbJzY0iaHmDIZs/f/ojkHtqHsn/AELpfuMMN2pbRhV3/AsZpNAadAvi5OCJ5pwGDdq2z+FY7MN2pbRhV3/AsWkhXcc0ZfxW2fwrHZhu1LaMKu/4Fi0kIByjAkALJfrZg8y/its/hWOzDdqW0YVd/wACxeTzmavlW2fwrHZhu1LaMKu94Fi8nnAg88r6AIePykQXbZZTgAASzDdqE56sm5oaJgWBDbr6lCrN/eBi8nnGGAUkxzIXcexJ24N2oyyQOpBAZ/nXiYvJ5w7RnLYj5QgDQkgq+PxXx+IqNp1ByZCMj4kAhtofi8d+F478Iot00c2Q/eFgNITMZTJ9Pj09VHC0kenqo4Wkj09VHC0kLvWQkIqBGZJLJdu0DlebLtMgGKBiBmCCXly8+XnSDuL7o5tVHC0kK+mqiarp4UkK3jwleSYROogzlzKqOFpIV9NVE1XTwpIVvFf6eD7sU4kSMOxHddByPUZga/KKYc99jqnnzKqOFpIV9NVE1XTwpIVpFatdVd9lQRw0r58yqjhaSFfTVRNV08KSFaRWrXVXfZWUcNK2fMqo4WkhX01UTVdPCkhWkVq11V32VlHDStnzKqOFpIV9NXE1XTwpIVpFatdVd9lZRw8rZ8yqjhaSEJ2pT7f/AFCxGCOoOqdlSCIgcN1wpIVpFaldVd9sRXlbPmCelv0AAuV5IoIAMunMIrIXdciFQC/QSbheSK8kUfOR9N2RyfUOhAOiDNdebZrXJ282yQ2ABGZyIDM5DrDILvAHUkm9X//aAAwDAQACAAMAAAAQ++4+++++++/4+++o+++++++rh+p6oHsvV0+++++vWoW/+s3+++++7WoW9+W++++++9co09+W++++++tMM8Nkc+++++++++++W+10842++++++W9Wp/8AKvvvvvvlvVqfwa/vvvvvlvXqf1Q/vvvvvlr37dmyfvvvvvvvvvvvvvv/xAAlEQEAAQMCBgMBAQAAAAAAAAABEQAhMVHwIEFhcaGxEDCRgcH/2gAIAQMBAT8QpcgXAuKIMOCdF5hp9BEbtSsnngTabKA2vrpPz4ivFfQCcMI6wiMdYbUORQiQRWAuFjOZul5EoOJGU1ZgO3P+Vc4tYzbnOOTlvXiqBLDMQvNPcNyl2A9U/WFJMFJhOZDJ+MlZHQHbi2zVrZda+Ko2vb4y2PerZqETHM0aMOBzGY4ts1a2XWviqNr2+Mtj3ryeMMNFg7jMfk02wOAtDnP614qg62RY6ns91I5cibt1enNrY968njAqkzQdjfdaKHcwpJ/rUcb0R6a2d/tOzQay3OzLRAJC6lWN37PGeqBRN6CKCM0FKFS835XXfjVhEcfjPVb3pTdNNYN1qPDBVc2jperqmRfkmYzjlE0oGaRjo3vcR/o8fjPVb3pTdNNYN1vg7B3e9bdrxvGeq3vSm6aawbrfB2Dv9q27XjBLqMocutDFhsMhNyTlJDpV61CUTszagzyNgLh0aX7sRCZjETUgXIiEzK0TM0CRhIHPJP3/AP/EACcRAQACAAQFBQEBAQAAAAAAAAEAESExYbEgQXGh8BAwUZHRgcHx/9oACAECAQE/EIItgvCNFiF418hydfYcZRLoK/o1yxBUxw3r179O++wNzjBOoibRrQFiNogLbVLi9KycOcOrpqH45rFUJwxYY6TvsGmFSrivAqgdDB1rWac4EFQ5Lk/eDMr4F68XcuxPO0jvvoR4Gj0fFei5XzhGG6VfF3LsTztI776EeBo9H3m7jPu749Eq/uoFhxb5mX5O+wRDDntHZlSWBdZYFf56PvN3GglMv/B+S9A0A2jto1B3n/PfkuXx8lY8sYyQteAozcj3O8bwkLWe6DFjBlDNF9zSfcBkePvG88rWeyd0zOrcgE2gD+xsebg1uq+M4WFiX+YZjx943nlaz2TumZ1bk83WZ/TsniPg4+8bzytZ7J3TM6tyebrNjsniPg4198DkLz0jGrrrJz/w3By2YAD1IgULZicyBFHPTWcZVgDdNVTG42wUxprI5+//AP/EACcQAQABAwMDBAMBAQAAAAAAAAERACFREDHwQEFhIHGBwTCRofFg/9oACAEBAAE/EOsNm2FQQG3hfuuUfdFpC71ZHyr89AI+2mxBTuFB3Bbek5diEqRGY2IMu9F+QUd0XeyRIAGJDpvV1QkPO45MwIzCe1Cq16kgFLKN0xYNhhuDoYjHdu8MzYRDpvV3BuJbhGyJIjZFGnrpdNO6JTxBgKEbkmhCDLkKBAJYEs6LVjvaMpNrCStpLPY8GF0c/qqZGJN4WDeR31uk12IMoWPnQF2Kd0sC8mmxaFy95Knb/dOT4FuC6FDETHeN9e3YRWFyF8hmiUxO25IznHZE/IXryO3S87U/iW9OwMSoT5p6nouoxJYIxJHcvrd4PHTlzuNPMYaAkBExkSk4koKBQSxIEsid6AWd/wAkkdYBPSv07vB46cudxp5jDVSllZ6ICcj9O7weOnLncaeYw1QmdSERgBtQkMgRFiV2u9Dkfp3eDx05crjTzGHQNGJLivpH6d3g8fRy75jDoGjAbm+UP6NGlj0QiBJsLZfNSn9XkGB9h1OgUyKD8a3RFoEJQxZvxQIo3Al30bbdOYw6BoCuGzIGSAcHl6brzPbdvEYAB8p2fRzGHQNJfGjvyC40xOECz8qK1/kVf5FTbaUSyBMeAPig6sizGFG54oMITKupYkSokMJIzNi8b207RjpZEvZ37v8AjM6nOpzqQSDsofIUXkUNAbqxQ5BuKT4K5H9UgOSZ8Hs0UCRBEdkYrl/1UnP/AJXN/qlg5uiH7OkzqarkLz2XqWhs7YMKZsQxOzCdHnU1XJXnstFsrtfcJfAjaKjzNbPXCBJCMwb1++SdiWMd2ztVvfuskEECRPBt0hTOpquSvPZaIebxTZ526bRgaZ1NVyV57LVDts87dNcwKM6mq5K89lqh22Oduj6QFGdTVKkE6yV57LVDtsc7dF0sKM6lTxGzsv1mh4sfIaA+EWgvYQAiIT2XTnstUJWxztWz81DCjf4FSwngBa5Z9Us1dVmwjJ4aFvNTJSIncSBN7CbQl2xMEe4DXKPquUfVEKzsqFxSDt3SaDbsiVeDSRgudvNbotEJCDt4adlkqUdi2Bp8c/olFSDNPgq5JoFu6nV//9k=" alt="Durlum logosu"></div>'
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
            "Adet/saat, m²/saat ve fire sınırlarını işlem bazında yönet.",
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


@st.cache_data(ttl=30, show_spinner=False)
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
    try:
        init_db()
    except Exception as exc:
        st.error(str(exc))
        st.warning(
            "Bu koruma özellikle açık bırakıldı: kalıcı veritabanı olmadan "
            "bulutta yeni kayıt kabul edilmeyecek."
        )
        st.stop()
    init_auth_state()

    is_manager = bool(st.session_state.manager_authenticated)

    render_durlum_header(is_manager)
    st.caption("Sürüm 2.17.1 · FAZLA ÜRETİM DESTEĞİ")
    st.markdown(
        """
        <div style="padding:10px 14px;border:1px solid #16a34a;border-radius:10px;
        background:#f0fdf4;font-weight:700;margin-bottom:10px;">
        ✓ AKTİF SÜRÜM 2.17.0 — GÜNLÜK HEDEF · ÇOKLU EKİP · DİĞER ÇALIŞMA
        </div>
        """,
        unsafe_allow_html=True,
    )
    render_database_status()
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
                <div class="sidebar-mark"><img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wgARCADIAMgDASIAAhEBAxEB/8QAHAABAAMBAQEBAQAAAAAAAAAAAAYHCAUEAwIB/8QAGwEBAAIDAQEAAAAAAAAAAAAAAAQFAgMGAQf/2gAMAwEAAhADEAAAAfCPqPw8TXTIhTSlHQbSOC0owAAAAAB78cvAl/h0So8JMMBYNfWDAtL2znozOfOdjFB2XzoAAAAAB3OH99e2VuT54VjyxY1KZQ3S1TfR7sduPct3EwznozOc2useNXDBteyDzviW1nhmC4Ks0lKgQDyzn21HQ5f5dg192/zQJMIAAABqXLWpea7Ll1zY1cw7G7M56Mzn5loODTmDQLTj21UttZ6s26SzbpKbXU5YtdVtMr5tX31+V7ywSIoAAADUuWtS812XLrmxq5h2N2Zz0ZnPzLQcGnMGgWnHtqpba2as26SzbpKZXU7VVq1VfcsFnSgAAAANO5ivGh6mVwC2nM9qznozOdlTaC4H9lFVewGfGGzNuks26Suecp2qrVqq+5YLOlAAAAAfT5iR/qNI8uS8PzM9fp7kaEmRlhs91lV9qCl6TOMWtSq7SjCdWAAAAAAAAAAdfUGX9Qcj39OVXalZXXN/B6fjZ0v4er4Pfw9P9898r1f08j0ef3EPfAAAAAOvqDL+oOR7+nJhD5hjn26Tuyk/cdCUDf1A6ZN1dTl+WqvO84v2x2dSF9Dpbo2Wh9G+PgAAAAdfUGX9Qcj39OTCHzDHPt0ndlJ+46EoG/qB0ybqi8oi8Sc6vK6u3TFp9AZ957lod58tAAAAA6+oMv6g5Hv6cmEPmGOfbpO7KT9x0JQN/UDpk3VF5RF4k51OX1NumMT2BT3z3LQ7z5aAAAAB19QZb0Dyvc1/LIH4tmm/Y4/dB1cjoG26ZtaO9Yv6o9Fm9fqRPo7NPjntXTJ7nkdt81AAAAAAAAAAAAA//8QAKRAAAQMDAwMEAwEBAAAAAAAABgQFNQACAwEwMxY0NhAUMUATFSARIv/aAAgBAQABBQL7gmjwrXPp1tohT40rxvo03uslmLJraoS2ZbP4CJeime30P/aO/KgyVrfrkcv4CJeime38Oa9Pl1Uo81KFv5Mfowjt7zWgQg/xsHErSoopnsIYgvxEg6laUAuyJ3jTolvrElsvd+iW+uiW+swMkutcm/I1q9sfwaJ2ZcrtQJGIiterqKZ5N2xvEAXxSfyOlTgmRa2X6ZLTeY22yNJYMFkKKZ5N2xvEAXxSfyOjznYYY3mNttjiWDBZCimeTdsbxAF8Un8jo857VeezTJkvy67bbHEsGCyFFM8m7Y3iAL4pP5HR5zbrZGksGCyFFM8m7Y3iAL4pP5HR5zbrLf8AkaHpNesag9tVIltFM8l7YjbMjs3CrMpadKT+R0ec26GOdmZD6lM8OOFi9r9U/kdHnNu2X3Y7tCJyt06jcq6jcqUKMirNgU5Ut/UjlXUjlXUjlSfXNmW+1JqfcTjjv+s0StHnP9ZolaPOf6zRK0ec9uDJfpqny6aW2636+2zVdZdZromy66e1zV7XNXtc1XYMlu60StHnOIwTzEiE7RlNtcZ/BKz4FbfttErR5ziME8xIhO0ZTbXGPTp+oRtDna7InNb+uQsbza9J3ON22iVo85xGCeYkQnaMptrjDSGC4YmggTtHOO22iVo85xGCeYkQnaMptrjDSGC4YmggTtHOO22iVo85xGCeYkQnaMptrjDSGC4YmggTtHOO22iVo85xGCeIkQnaMptrjDSGC4YmggTtHOO22iVo85w6/S5kvs0yWIh1E3qaMptrjDSFC4YnggTtHOO223Jbiceom2jFenX5hp/0aL7SVtv06iba6ibaKFWJY7N7834kBU7o1rUKuyNE1v70hUtAe5pUKZe/N+RD9r//xAA3EQABAwEDCQUGBwEAAAAAAAABAAIDBAURcQYSITE0NYGxwRMgM1FyFDAyQWHSECJCU4KR0RX/2gAIAQMBAT8BVPCaiVsTfmq6yZaCMSPcDpu9w1pec0J1M4C8dfxszbYsVlHsrfV0PuInBjrynTi68kcPoPno/CxqJlZORL8IQgs+nnZGGgPOpZR7K31dCq6gpY6N72xi+5WVQ001Gx8kYJ08yrFpIJxL2rAbiv8AnUDyWZgVbC2CofE3UD3smvEkwCrd7wYf6so9lb6uhVo7BJgrF2CPjzKyf+GbFS14s+0pnlt993IKqm9pndLddf3smvEkwCrd7wYf6so9lb6uhVo7BJgrF2CPjzKyf+GbFWzt8nDkO/k9M2OpLHfqCloWS1LKknS1ZR7K31dCgI62mu/S4Kmp2UsQhj1BZP8AwzYq2dvk4ch39S9rqP3D/ZT55ZRc9xPFMnliFzHEcV7ZU/uH+yrFoxVMe4vcNPyNytKLsKt8YJN3nr1e8yb8KTFWwCa+S76cgs0otI1hZjvJZjvJEEa+/k34UmKpt9TYfarU3jTYjmsodj4jqhI2GnEj9QCE0Zi7YH8t1/BS9jW0pI0tI7+TfhSYqm31Nh9qtTeNNiOayh2PiOqrN3O9PRQ7p/h0Vm7tbgevfyb8KTFU2+psPtVqbxpsRzWUOx8R1VZu53p6KHdP8Ois3drcD17+T00cUT89wGlNrIoLYfI4/lOi/gE6WilIe5zSRq0hW9PFJSXMcDp81V1EJoHNDxfm+aiqIRZeZni/M8/orPqIW2c1rni+4/PH3/8A/8QAMxEAAAQEAQkGBwEAAAAAAAAAAAECAwQFEXE0EhMgITEzNUHBFDBRgdHwECIyUoKR4bH/2gAIAQIBAT8BDzhMtm4fIQswRFryEl3BmSSqYJ8jOnxjsMuwk2/VbqXcLTlJoQJk/A/hMolUM18m0xnYt5pS6maeYk2/VbqQhYt9cSlKlnSomEU+3EqShdC/gmcQ60aMhVNQ7ZFp+bLMQzhuspWraelO/oQIbh7t/QSbfqt1IQeLRcTPFr8v8E32t2DcIcXBNpI6U/oh2sy0lvw0p39CBDcPdv6CTb9VupCDxaLiZ4tfl/gm+1uwluER756c4bNTJKLkG4pTbKmCLUYk2/VbqQM1wz9eZGHnlPuG4vaYm+1uwluER759x2dn7C/QS02g6oSRBTTa9akkY7Oz9hfoTOJNhaSJJHchBOZ2HSulLd5Ot4iwlx0hEe+YqQqRjKLxGUnxFa6c63iLB7hjd/UQGDf98hJ8T5A0KcdNCdpmM0sl5umvYEZyGfIthkenOt4iwe4Y3f1EBg3/AHyEnxPkIfGJuHMf+XURuNVctOdbxFg9wxu/qIDBv++Qk+J8hD4xNw5xD8uojcaq5ac4bWtackqg4Zx2XJQRay/oSiJQRpSRlW4lLTiIiqk01CHZcKLSZpPaFsuduysk6ZXURbLhxhmST29//8QAQRAAAQICBgYFCwIEBwAAAAAAAgABAwQQERJyc7ETITBxssExNFGBgxQgIzIzQEFhdIKSIuEFFZGhJEJSYmPR8P/aAAgBAQAGPwL3woceG0QNG71PvZdUhqYhQhYIY1VC25vcHrKxDFrRn2MmeW/hNuC/QcQSJyUUghPLR4OuJAfs7W808J82omu7hb3Cehj69hi3sz6//fJSITDE/oGZzhn6ut+lqlOxysWBgkz2CtD6tltf9PNPCfNqJru4W9wGJDKyY9Dq1ElTA/joYlQv3OzrQwobQIFddltbu/zekjItFAF6rXxfcvXjv9zf9J40Eojk42f1PRNd3CyAneNW7M/rIY0F4lp4jD+p96mdO5+js1WH3rpjfl+yCWevRvH0fzqtVLpjfl+y6Y35fsvRxooF86nRy8TW7fFvi20lBb4ha/rrUWYNqxBq6mUQNFoTDXVXXW1E13cLKFdZBjNk6nfs50Q/q246BaPGCE5dFp0xC7EL62dkOE2b7SVwhyU3d5qYwudE13cLKFdZBjNk6nfs50Q/q246JS6Sk8NkGE2b7SVwhyU3d5qYwudE13cLKFdZBjNk6nfs50Q/q246JS6SqGNEZm+DEqzJzftJ9pK4Q5Kbu81MYXOia7uFlCusgxmydTv2c6If1bcdEpdLltpXCHJTd3mpjC50TXdwsoV1kGM2Tqd+znRD+rbjolLpcttJv/xDkpmDD1mQ6mUYo8A4QvDqZyb50TXdwsoVxlooTtpGJia0pnyhha3ZqsvX20Q/q246JS6XLbeSEXpYXQ3aPmTXdwsoNT+khjYNvMh/Vtx0Sl0uW2YgJxJuh2VXlcRdbNdbNFFilbiF0k6twYhQy7RepdbNdbP+y62f9lDeG/8AiCiNZf8A3Vr2j/mKhfzErRVPY1s/u8ljBnRKXS93ksYM6JS6Xu8ljBnRKXSVYwyJu1mVbwzZrqqFnJ+xl7I/xVRM4v8ANVtCN2ur2R/ivZH+K9if4rXDJt7bWSxgzolLpKDvLNTuCeSg7iyoK4KlMIcvNjRWhsMeGNtjZuzaSWMGdEpdJQd5ZqdwTyUHcWVBXBUphDktPo9L+phs11IY4g4a6nF1FmLGksN6tdSKI0PREL1ONdamsIstpJYwZ0Sl0lB3lmp3BPJQdxZUFcFSmCGS8Rl4jqbutmpq+2SmsIstpJYwZ0Sl0lB3lmp3BPJQdxZUFcFSmCGS8Rl4jqb3Nmpq+2SmsIstpJYwZ0Sl0lB3lmp3BPJQdxZUFcFSmCGS8Rl4jqb3Nmpq+2SmsIstpJYwZ0Sl0lB3lmp3BPJQdxZUFcFSmCGS8Rl4jqb3Nmpm+2SmsIstpJYwZ0Sl0kDf6SJkQE1Yk1TsmjwYbjEbo/U70FcFSmCGS8Rl4jqb3Nmymb7ZKawiy2kqZvZEYou79663DUs8vFGKws9dSOHFZylz16v8rqvyse+tl1uGutw0USCbRAstrZSwFNAxDDFnbuWjgzAxDts9TLRx44wztu9TqZhQpkDiEzVC29TAx4wwnctVamAGaByKGTM3d73/AP/EACkQAAEEAAQFBQEBAQAAAAAAAAEAEVHwECExoSAwQWHxQHGBkdGxweH/2gAIAQEAAT8h9YLRul3/AOi8EUQ2nSBnegbT3tnb7dToB3ITNyOJaXBA+gtRnTt0ZuYZw4L6vxrd36CBrggxAn9IEA4B1j0XHz7rNpSPgLWTk8Zbu/QQ4AxwIlDAX1A+CmsAZv3LrHbtiezVDHLH6QQOZKByVMmQxIMdsN2o75MYY9keUMGyGIKOyKAHQg1zdOy8HTJKImKDwdPB0LI8Wv0w/qbJzY0iaHmDIZs/f/ojkHtqHsn/AELpfuMMN2pbRhV3/AsZpNAadAvi5OCJ5pwGDdq2z+FY7MN2pbRhV3/AsWkhXcc0ZfxW2fwrHZhu1LaMKu/4Fi0kIByjAkALJfrZg8y/its/hWOzDdqW0YVd/wACxeTzmavlW2fwrHZhu1LaMKu94Fi8nnAg88r6AIePykQXbZZTgAASzDdqE56sm5oaJgWBDbr6lCrN/eBi8nnGGAUkxzIXcexJ24N2oyyQOpBAZ/nXiYvJ5w7RnLYj5QgDQkgq+PxXx+IqNp1ByZCMj4kAhtofi8d+F478Iot00c2Q/eFgNITMZTJ9Pj09VHC0kenqo4Wkj09VHC0kLvWQkIqBGZJLJdu0DlebLtMgGKBiBmCCXly8+XnSDuL7o5tVHC0kK+mqiarp4UkK3jwleSYROogzlzKqOFpIV9NVE1XTwpIVvFf6eD7sU4kSMOxHddByPUZga/KKYc99jqnnzKqOFpIV9NVE1XTwpIVpFatdVd9lQRw0r58yqjhaSFfTVRNV08KSFaRWrXVXfZWUcNK2fMqo4WkhX01UTVdPCkhWkVq11V32VlHDStnzKqOFpIV9NXE1XTwpIVpFatdVd9lZRw8rZ8yqjhaSEJ2pT7f/AFCxGCOoOqdlSCIgcN1wpIVpFaldVd9sRXlbPmCelv0AAuV5IoIAMunMIrIXdciFQC/QSbheSK8kUfOR9N2RyfUOhAOiDNdebZrXJ282yQ2ABGZyIDM5DrDILvAHUkm9X//aAAwDAQACAAMAAAAQ++4+++++++/4+++o+++++++rh+p6oHsvV0+++++vWoW/+s3+++++7WoW9+W++++++9co09+W++++++tMM8Nkc+++++++++++W+10842++++++W9Wp/8AKvvvvvvlvVqfwa/vvvvvlvXqf1Q/vvvvvlr37dmyfvvvvvvvvvvvvvv/xAAlEQEAAQMCBgMBAQAAAAAAAAABEQAhMVHwIEFhcaGxEDCRgcH/2gAIAQMBAT8QpcgXAuKIMOCdF5hp9BEbtSsnngTabKA2vrpPz4ivFfQCcMI6wiMdYbUORQiQRWAuFjOZul5EoOJGU1ZgO3P+Vc4tYzbnOOTlvXiqBLDMQvNPcNyl2A9U/WFJMFJhOZDJ+MlZHQHbi2zVrZda+Ko2vb4y2PerZqETHM0aMOBzGY4ts1a2XWviqNr2+Mtj3ryeMMNFg7jMfk02wOAtDnP614qg62RY6ns91I5cibt1enNrY968njAqkzQdjfdaKHcwpJ/rUcb0R6a2d/tOzQay3OzLRAJC6lWN37PGeqBRN6CKCM0FKFS835XXfjVhEcfjPVb3pTdNNYN1qPDBVc2jperqmRfkmYzjlE0oGaRjo3vcR/o8fjPVb3pTdNNYN1vg7B3e9bdrxvGeq3vSm6aawbrfB2Dv9q27XjBLqMocutDFhsMhNyTlJDpV61CUTszagzyNgLh0aX7sRCZjETUgXIiEzK0TM0CRhIHPJP3/AP/EACcRAQACAAQFBQEBAQAAAAAAAAEAESExYbEgQXGh8BAwUZHRgcHx/9oACAECAQE/EIItgvCNFiF418hydfYcZRLoK/o1yxBUxw3r179O++wNzjBOoibRrQFiNogLbVLi9KycOcOrpqH45rFUJwxYY6TvsGmFSrivAqgdDB1rWac4EFQ5Lk/eDMr4F68XcuxPO0jvvoR4Gj0fFei5XzhGG6VfF3LsTztI776EeBo9H3m7jPu749Eq/uoFhxb5mX5O+wRDDntHZlSWBdZYFf56PvN3GglMv/B+S9A0A2jto1B3n/PfkuXx8lY8sYyQteAozcj3O8bwkLWe6DFjBlDNF9zSfcBkePvG88rWeyd0zOrcgE2gD+xsebg1uq+M4WFiX+YZjx943nlaz2TumZ1bk83WZ/TsniPg4+8bzytZ7J3TM6tyebrNjsniPg4198DkLz0jGrrrJz/w3By2YAD1IgULZicyBFHPTWcZVgDdNVTG42wUxprI5+//AP/EACcQAQABAwMDBAMBAQAAAAAAAAERACFREDHwQEFhIHGBwTCRofFg/9oACAEBAAE/EOsNm2FQQG3hfuuUfdFpC71ZHyr89AI+2mxBTuFB3Bbek5diEqRGY2IMu9F+QUd0XeyRIAGJDpvV1QkPO45MwIzCe1Cq16kgFLKN0xYNhhuDoYjHdu8MzYRDpvV3BuJbhGyJIjZFGnrpdNO6JTxBgKEbkmhCDLkKBAJYEs6LVjvaMpNrCStpLPY8GF0c/qqZGJN4WDeR31uk12IMoWPnQF2Kd0sC8mmxaFy95Knb/dOT4FuC6FDETHeN9e3YRWFyF8hmiUxO25IznHZE/IXryO3S87U/iW9OwMSoT5p6nouoxJYIxJHcvrd4PHTlzuNPMYaAkBExkSk4koKBQSxIEsid6AWd/wAkkdYBPSv07vB46cudxp5jDVSllZ6ICcj9O7weOnLncaeYw1QmdSERgBtQkMgRFiV2u9Dkfp3eDx05crjTzGHQNGJLivpH6d3g8fRy75jDoGjAbm+UP6NGlj0QiBJsLZfNSn9XkGB9h1OgUyKD8a3RFoEJQxZvxQIo3Al30bbdOYw6BoCuGzIGSAcHl6brzPbdvEYAB8p2fRzGHQNJfGjvyC40xOECz8qK1/kVf5FTbaUSyBMeAPig6sizGFG54oMITKupYkSokMJIzNi8b207RjpZEvZ37v8AjM6nOpzqQSDsofIUXkUNAbqxQ5BuKT4K5H9UgOSZ8Hs0UCRBEdkYrl/1UnP/AJXN/qlg5uiH7OkzqarkLz2XqWhs7YMKZsQxOzCdHnU1XJXnstFsrtfcJfAjaKjzNbPXCBJCMwb1++SdiWMd2ztVvfuskEECRPBt0hTOpquSvPZaIebxTZ526bRgaZ1NVyV57LVDts87dNcwKM6mq5K89lqh22Oduj6QFGdTVKkE6yV57LVDtsc7dF0sKM6lTxGzsv1mh4sfIaA+EWgvYQAiIT2XTnstUJWxztWz81DCjf4FSwngBa5Z9Us1dVmwjJ4aFvNTJSIncSBN7CbQl2xMEe4DXKPquUfVEKzsqFxSDt3SaDbsiVeDSRgudvNbotEJCDt4adlkqUdi2Bp8c/olFSDNPgq5JoFu6nV//9k=" alt="Durlum logosu"></div>
                <div>
                    <div class="sidebar-brand-title">DURLUM FLOW</div>
                    <div class="sidebar-brand-subtitle">Production Intelligence System</div>
                </div>
            </div>
            <div class="sidebar-version">
                Sürüm 2.17.1 · FAZLA ÜRETİM DESTEĞİ
            </div>
            """,
            unsafe_allow_html=True,
        )
        manager_login_panel()
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
        # Yetki değiştiğinde eski ve artık izin verilmeyen sayfa seçimini temizle.
        current_selected_page = st.session_state.get("main_page_selector")
        if current_selected_page not in allowed_pages:
            st.session_state["main_page_selector"] = allowed_pages[0]

        page = st.radio(
            "Sayfa",
            allowed_pages,
            index=0,
            key="main_page_selector",
            on_change=handle_main_page_change,
            format_func=lambda option: f"{page_icons.get(option, '•')}  {option}",
        )

        uploaded_excel = None
        if is_manager:
            st.divider()
            st.markdown("**Yönetici araçları**")

            if st.toggle(
                "Bildirim merkezini aç",
                value=False,
                key="lazy_notification_center",
                help="Ağır yönetici analizleri yalnızca bu seçenek açıldığında çalışır.",
            ):
                render_notification_center()

            if st.toggle(
                "Üretim çıktısı yönetimini aç",
                value=False,
                key="lazy_output_manager",
                help="PDF/Excel yükleme ve OC silme alanını gerektiğinde aç.",
            ):
                sidebar_production_output_manager()

            if st.toggle(
                "Veri yedekleme araçlarını aç",
                value=False,
                key="lazy_backup_tools",
            ):
                render_manager_backup_download()


            if st.toggle(
                "Ayar Excel'i yükle",
                value=False,
                key="lazy_config_upload",
            ):
                uploaded_excel = st.file_uploader(
                    "Ayar Excel'ini değiştir",
                    type=["xlsx"],
                    help="Boş bırakırsan uygulama klasöründeki ayar_dosyasi.xlsx kullanılır.",
                )

    (reasons, combinations), source_info = load_config(uploaded_excel)
    if is_manager:
        st.sidebar.caption(source_info)

    # Önceki sayfanın widget'larını anında kaldır; ağır sorgu sürerken iki ekran üst üste görünmesin.
    page_root = st.empty()
    page_root.empty()
    with page_root.container():
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


def render_other_work_entry(reasons):
    """POS üretimi dışındaki günlük işleri saat bilgisiyle kaydeder."""
    version = int(st.session_state.get("other_work_form_version", 0))
    success = st.session_state.pop("other_work_success", None)
    if success:
        st.success(success)

    st.markdown("### Diğer Çalışma Kaydı")
    st.caption(
        "Toplantı, bakım, sevkiyat desteği, düzenleme ve benzeri POS dışı işleri "
        "buradan manuel olarak kaydedebilirsin."
    )
    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            work_date = st.date_input(
                "Tarih", value=date.today(), key=f"other_date_{version}"
            )
        with c2:
            worker = st.selectbox(
                "Çalışan", WORKER_NAMES, key=f"other_worker_{version}"
            )
        with c3:
            declared_type = st.selectbox(
                "Çalışma şekli",
                ["Tam zamanlı", "Yarı zamanlı", "Mesaili"],
                key=f"other_type_{version}",
            )

        description = st.text_input(
            "Ne yapıldı?",
            key=f"other_description_{version}",
            placeholder="Örnek: Sevkiyat hazırlığı ve depo düzenlemesi",
        )
        d1, d2 = st.columns(2)
        with d1:
            hours = st.number_input(
                "Çalışma süresi (saat)",
                min_value=0.25,
                max_value=24.0,
                value=1.0,
                step=0.25,
                key=f"other_hours_{version}",
            )
        with d2:
            reason = st.selectbox(
                "Neden / açıklama türü",
                [""] + list(reasons),
                disabled=(declared_type != "Yarı zamanlı"),
                key=f"other_reason_{version}",
                format_func=lambda value: "Seçin" if value == "" else value,
            )
        team_members = st.multiselect(
            "Çalışma ekibi (isteğe bağlı)",
            [name for name in WORKER_NAMES if name],
            key=f"other_team_{version}",
            help=(
                "Bu işe ana çalışan dışında katılan kişileri seçin. "
                "Seçilen kişiler yönetici çalışan özetlerinde de görünür."
            ),
        )
        note = st.text_area(
            "Not", key=f"other_note_{version}",
            placeholder="Varsa ek bilgi yazın.",
        )

        existing_hours = get_registered_daily_hours(worker, work_date) if worker else 0.0
        final_type = determine_work_type(
            work_date, declared_type, float(hours), existing_hours
        )
        st.info(
            f"Bu kayıtla günlük toplam süre: {existing_hours + float(hours):.2f} saat · "
            f"Sisteme **{final_type}** olarak yazılacak."
        )

        if st.button(
            "Diğer Çalışmayı Kaydet",
            type="primary",
            use_container_width=True,
            key=f"save_other_work_{version}",
        ):
            errors = []
            if not worker:
                errors.append("Çalışan seçmelisin.")
            if not str(description).strip():
                errors.append("Yapılan işi yazmalısın.")
            if float(existing_hours) + float(hours) > 24.0001:
                errors.append("Günlük toplam çalışma süresi 24 saati aşamaz.")
            if declared_type == "Yarı zamanlı" and not reason:
                errors.append("Yarı zamanlı çalışma için neden seçmelisin.")
            if errors:
                for error in errors:
                    st.error(error)
            else:
                record_id = save_other_work_log({
                    "tarih": str(work_date),
                    "operator_ismi": worker,
                    "calisma_tipi": final_type,
                    "calisma_saati": float(hours),
                    "is_aciklamasi": description,
                    "participants_text": _join_participants(
                        [name for name in team_members if name != worker]
                    ),
                    "neden": reason,
                    "notlar": note,
                })
                st.session_state.other_work_form_version = version + 1
                st.session_state.other_work_success = (
                    f"Diğer çalışma kaydedildi. Kayıt no: {record_id}."
                )
                st.rerun()


def yeni_kayit_page(reasons, combinations, is_manager: bool):
    st.subheader("Yeni Operasyon Kaydı")
    st.caption(
        "Kayıt beş kısa adımda hazırlanır. Son adımda bütün bilgiler kontrol edilmeden "
        "veritabanına yazılmaz."
    )

    success_message = st.session_state.pop("operation_entry_success", None)
    if success_message:
        st.success(success_message)

    entry_mode = st.radio(
        "Kayıt türü",
        ["POS / Operasyon", "Diğer Çalışma"],
        horizontal=True,
        key="new_entry_mode",
    )
    if entry_mode == "Diğer Çalışma":
        render_other_work_entry(reasons)
        return

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
        current_entries = wizard_data.get("entries", [])
        return _entries_total_hours(current_entries)


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

        saved_entry_groups = {}
        for saved_entry in wizard_data.get("entries", []):
            key = (int(saved_entry["item_id"]), str(saved_entry["operation_name"]))
            saved_entry_groups.setdefault(key, []).append(saved_entry)

        entries = []
        for _, item in selected_items.iterrows():
            item_id = int(item["item_id"])
            requested = int(item["requested_qty"])
            for operation in selected_map.get(item_id, []):
                progress_row = lookup.get((item_id, operation))
                completed = int(progress_row["completed_qty"]) if progress_row is not None else 0
                remaining = max(requested - completed, 0)
                normalized_operation = _normalize_header(operation)
                operation_kind = _operation_kind(operation)
                saved_entries = saved_entry_groups.get((item_id, operation), [])

                if operation_kind == "laser_cut_out":
                    processed_key = (
                        f"processed_{item_id}_{normalized_operation}_{version}"
                    )
                    saved_entry = saved_entries[0] if saved_entries else {}
                    processed = int(
                        st.session_state.get(
                            processed_key,
                            saved_entry.get("processed_qty", 0),
                        )
                        or 0
                    )
                    if processed > 0:
                        entries.append({
                            "item_id": item_id,
                            "pos": str(item["pos"]),
                            "operation_name": operation,
                            "processed_qty": processed,
                            "fire_qty": 0,
                            "good_qty": processed,
                            "remaining_before": remaining,
                            "unit_area_mm2": float(item["unit_area_mm2"]),
                            "boy_mm": float(item["boy_mm"]),
                            "en_mm": float(item["en_mm"]),
                            "operation_hours": 0.0,
                            "laser_plate_qty": 0,
                            "material_type": "",
                            "thickness_mm": 0.0,
                            "laser_lot_no": 0,
                            "abkant_work_mode": "",
                            "abkant_coworker": "",
                            "abkant_team_qty": 0,
                            "abkant_teams_per_piece": 1,
                            "abkant_team_excess": 0,
                            "abkant_long_bend_qty": 0,
                            "abkant_short_bend_qty": 0,
                            "abkant_long_single_bend_qty": 0,
                            "abkant_long_double_bend_qty": 0,
                            "abkant_short_single_bend_qty": 0,
                            "abkant_short_double_bend_qty": 0,
                            "abkant_manual_override": False,
                            "participants_text": "",
                            "piece_weight_kg": 0.0,
                            "bend_type": "",
                            # Laser Cut Out yalnızca adet girer; alan POS ölçüsünden otomatik hesaplanır.
                            "calculated_area_mm2": processed * float(item["unit_area_mm2"]),
                        })
                    continue

                if operation_kind == "laser":
                    lot_count_key = f"laserlotcount_{item_id}_{normalized_operation}_{version}"
                    default_lot_count = max(len(saved_entries), 1)
                    lot_count = int(
                        st.session_state.get(lot_count_key, default_lot_count) or 1
                    )
                    lot_remaining = remaining

                    for lot_index in range(1, lot_count + 1):
                        saved_lot = (
                            saved_entries[lot_index - 1]
                            if lot_index - 1 < len(saved_entries)
                            else {}
                        )
                        suffix = f"laserlot_{lot_index}_{version}"
                        processed_key = f"processed_{item_id}_{normalized_operation}_{suffix}"
                        fire_key = f"opfire_{item_id}_{normalized_operation}_{suffix}"
                        hours_key = f"ophours_{item_id}_{normalized_operation}_{suffix}"
                        plate_key = f"laserplates_{item_id}_{normalized_operation}_{suffix}"
                        material_key = f"material_{item_id}_{normalized_operation}_{suffix}"
                        thickness_key = f"thickness_{item_id}_{normalized_operation}_{suffix}"
                        long_qty_key = f"laserlongqty_{item_id}_{normalized_operation}_{suffix}"
                        short_qty_key = f"lasershortqty_{item_id}_{normalized_operation}_{suffix}"
                        long_per_team_key = f"laserlongteam_{item_id}_{normalized_operation}_{suffix}"
                        short_per_team_key = f"lasershortteam_{item_id}_{normalized_operation}_{suffix}"

                        long_edge_qty = int(
                            st.session_state.get(
                                long_qty_key,
                                saved_lot.get("laser_long_edge_qty", 0),
                            ) or 0
                        )
                        short_edge_qty = int(
                            st.session_state.get(
                                short_qty_key,
                                saved_lot.get("laser_short_edge_qty", 0),
                            ) or 0
                        )
                        long_per_team = max(int(
                            st.session_state.get(
                                long_per_team_key,
                                saved_lot.get("laser_long_per_team", 2),
                            ) or 2
                        ), 1)
                        short_per_team = max(int(
                            st.session_state.get(
                                short_per_team_key,
                                saved_lot.get("laser_short_per_team", 2),
                            ) or 2
                        ), 1)
                        processed = min(
                            long_edge_qty // long_per_team,
                            short_edge_qty // short_per_team,
                        )
                        fire = int(
                            st.session_state.get(
                                fire_key,
                                saved_lot.get("fire_qty", 0),
                            )
                            or 0
                        )
                        operation_hours = float(
                            st.session_state.get(
                                hours_key,
                                saved_lot.get("operation_hours", 0.0),
                            )
                            or 0.0
                        )
                        good_qty = processed - fire

                        if processed > 0:
                            entries.append({
                                "item_id": item_id,
                                "pos": str(item["pos"]),
                                "operation_name": operation,
                                "processed_qty": processed,
                                "fire_qty": fire,
                                "good_qty": good_qty,
                                "remaining_before": lot_remaining,
                                "unit_area_mm2": float(item["unit_area_mm2"]),
                                "boy_mm": float(item["boy_mm"]),
                                "en_mm": float(item["en_mm"]),
                                "operation_hours": operation_hours,
                                "laser_plate_qty": int(
                                    st.session_state.get(
                                        plate_key,
                                        saved_lot.get("laser_plate_qty", 0),
                                    )
                                    or 0
                                ),
                                "material_type": str(
                                    st.session_state.get(
                                        material_key,
                                        saved_lot.get("material_type", ""),
                                    )
                                    or ""
                                ),
                                "thickness_mm": float(
                                    st.session_state.get(
                                        thickness_key,
                                        saved_lot.get("thickness_mm", 0.0),
                                    )
                                    or 0.0
                                ),
                                "laser_long_edge_qty": long_edge_qty,
                                "laser_short_edge_qty": short_edge_qty,
                                "laser_long_per_team": long_per_team,
                                "laser_short_per_team": short_per_team,
                                "laser_long_excess_qty": long_edge_qty - processed * long_per_team,
                                "laser_short_excess_qty": short_edge_qty - processed * short_per_team,
                                "laser_lot_no": lot_index,
                                "abkant_work_mode": "",
                                "abkant_coworker": "",
                                "abkant_team_qty": 0,
                                "abkant_teams_per_piece": 1,
                                "abkant_team_excess": 0,
                                "abkant_long_bend_qty": 0,
                                "abkant_short_bend_qty": 0,
                                "abkant_manual_override": False,
                                "participants_text": "",
                                "piece_weight_kg": 0.0,
                                "bend_type": "",
                                "calculated_area_mm2": max(good_qty, 0)
                                * float(item["unit_area_mm2"]),
                            })
                            lot_remaining = max(lot_remaining - max(good_qty, 0), 0)
                    continue

                processed_key = f"processed_{item_id}_{normalized_operation}_{version}"
                fire_key = f"opfire_{item_id}_{normalized_operation}_{version}"
                hours_key = f"ophours_{item_id}_{normalized_operation}_{version}"
                work_mode_key = f"abkantmode_{item_id}_{normalized_operation}_{version}"
                coworker_key = f"abkantcoworker_{item_id}_{normalized_operation}_{version}"
                abkant_team_qty_key = (
                    f"abkantteamqty_{item_id}_{normalized_operation}_{version}"
                )
                abkant_teams_per_piece_key = (
                    f"abkantteamsperpiece_{item_id}_{normalized_operation}_{version}"
                )
                abkant_long_bend_qty_key = (
                    f"abkantlongbendqty_{item_id}_{normalized_operation}_{version}"
                )
                abkant_short_bend_qty_key = (
                    f"abkantshortbendqty_{item_id}_{normalized_operation}_{version}"
                )
                abkant_long_single_bend_qty_key = (
                    f"abkantlongsinglebendqty_{item_id}_{normalized_operation}_{version}"
                )
                abkant_long_double_bend_qty_key = (
                    f"abkantlongdoublebendqty_{item_id}_{normalized_operation}_{version}"
                )
                abkant_short_single_bend_qty_key = (
                    f"abkantshortsinglebendqty_{item_id}_{normalized_operation}_{version}"
                )
                abkant_short_double_bend_qty_key = (
                    f"abkantshortdoublebendqty_{item_id}_{normalized_operation}_{version}"
                )
                abkant_manual_override_key = (
                    f"abkantmanualoverride_{item_id}_{normalized_operation}_{version}"
                )
                participants_key = f"participants_{item_id}_{normalized_operation}_{version}"
                weight_key = f"pieceweight_{item_id}_{normalized_operation}_{version}"
                bend_key = f"bendtype_{item_id}_{normalized_operation}_{version}"
                saved_entry = saved_entries[0] if saved_entries else {}

                saved_teams_per_piece = max(
                    int(saved_entry.get("abkant_teams_per_piece", 1) or 1),
                    1,
                )
                if operation_kind == "abkant":
                    # Abkantta günlük olarak tamamlanmış ürün adedi girmek zorunlu değildir.
                    # Kullanıcı yalnızca o gün yaptığı uzun/kısa kenar ve tek/çift büküm
                    # miktarını girer. Tamamlanan ürün adedi kayıt sırasında, geçmiş
                    # kenar üretimleriyle birlikte otomatik hesaplanır.
                    abkant_teams_per_piece = 1
                    abkant_team_qty = 0
                    processed = 0
                else:
                    abkant_teams_per_piece = 1
                    abkant_team_qty = 0
                    processed = int(
                        st.session_state.get(
                            processed_key,
                            saved_entry.get("processed_qty", 0),
                        )
                        or 0
                    )
                fire = int(
                    st.session_state.get(
                        fire_key,
                        saved_entry.get("fire_qty", 0),
                    )
                    or 0
                )
                operation_hours = float(
                    st.session_state.get(
                        hours_key,
                        saved_entry.get("operation_hours", 0.0),
                    )
                    or 0.0
                )
                abkant_activity = False
                if operation_kind == "abkant":
                    # Abkant satırı yalnızca gerçekten bir kenar/büküm adedi
                    # girildiğinde aktif sayılır. Süre alanının dolu olması tek
                    # başına boş POS satırlarını kayda dahil etmez.
                    abkant_edge_total_for_activity = sum(
                        max(
                            int(st.session_state.get(widget_key, 0) or 0),
                            0,
                        )
                        for widget_key in (
                            abkant_long_single_bend_qty_key,
                            abkant_long_double_bend_qty_key,
                            abkant_short_single_bend_qty_key,
                            abkant_short_double_bend_qty_key,
                        )
                    )
                    abkant_activity = abkant_edge_total_for_activity > 0

                if processed > 0 or abkant_activity:
                    entries.append({
                        "item_id": item_id,
                        "pos": str(item["pos"]),
                        "operation_name": operation,
                        "processed_qty": processed,
                        "fire_qty": fire,
                        "good_qty": processed - fire,
                        "remaining_before": remaining,
                        "unit_area_mm2": float(item["unit_area_mm2"]),
                        "boy_mm": float(item["boy_mm"]),
                        "en_mm": float(item["en_mm"]),
                        "operation_hours": operation_hours,
                        "laser_plate_qty": 0,
                        "material_type": "",
                        "thickness_mm": 0.0,
                        "laser_lot_no": 0,
                        "abkant_work_mode": str(
                            st.session_state.get(
                                work_mode_key,
                                saved_entry.get("abkant_work_mode", ""),
                            )
                        ) if operation_kind == "abkant" else "",
                        "abkant_coworker": (
                            str(
                                st.session_state.get(
                                    coworker_key,
                                    saved_entry.get("abkant_coworker", ""),
                                )
                            )
                            if operation_kind == "abkant"
                            and str(
                                st.session_state.get(
                                    work_mode_key,
                                    saved_entry.get("abkant_work_mode", ""),
                                )
                            ) == "Biriyle beraber çalıştı"
                            else ""
                        ),
                        "abkant_team_qty": (
                            abkant_team_qty if operation_kind == "abkant" else 0
                        ),
                        "abkant_teams_per_piece": (
                            abkant_teams_per_piece
                            if operation_kind == "abkant"
                            else 1
                        ),
                        "abkant_team_excess": (
                            abkant_team_qty % abkant_teams_per_piece
                            if operation_kind == "abkant"
                            else 0
                        ),
                        "abkant_long_single_bend_qty": (
                            int(st.session_state.get(
                                abkant_long_single_bend_qty_key,
                                saved_entry.get("abkant_long_single_bend_qty", 0),
                            ) or 0) if operation_kind == "abkant" else 0
                        ),
                        "abkant_long_double_bend_qty": (
                            int(st.session_state.get(
                                abkant_long_double_bend_qty_key,
                                saved_entry.get("abkant_long_double_bend_qty", 0),
                            ) or 0) if operation_kind == "abkant" else 0
                        ),
                        "abkant_short_single_bend_qty": (
                            int(st.session_state.get(
                                abkant_short_single_bend_qty_key,
                                saved_entry.get("abkant_short_single_bend_qty", 0),
                            ) or 0) if operation_kind == "abkant" else 0
                        ),
                        "abkant_short_double_bend_qty": (
                            int(st.session_state.get(
                                abkant_short_double_bend_qty_key,
                                saved_entry.get("abkant_short_double_bend_qty", 0),
                            ) or 0) if operation_kind == "abkant" else 0
                        ),
                        "abkant_long_bend_qty": (
                            int(st.session_state.get(abkant_long_single_bend_qty_key, 0) or 0)
                            + int(st.session_state.get(abkant_long_double_bend_qty_key, 0) or 0)
                            if operation_kind == "abkant" else 0
                        ),
                        "abkant_short_bend_qty": (
                            int(st.session_state.get(abkant_short_single_bend_qty_key, 0) or 0)
                            + int(st.session_state.get(abkant_short_double_bend_qty_key, 0) or 0)
                            if operation_kind == "abkant" else 0
                        ),
                        "abkant_manual_override": (
                            bool(
                                st.session_state.get(
                                    abkant_manual_override_key,
                                    saved_entry.get("abkant_manual_override", False),
                                )
                            )
                            if operation_kind == "abkant"
                            else False
                        ),
                        "participants_text": (
                            _join_participants(
                                st.session_state.get(
                                    participants_key,
                                    _split_participants(
                                        saved_entry.get("participants_text", "")
                                    ),
                                )
                            )
                            if operation_kind in {"boya", "paketleme", "sevkiyat"}
                            else ""
                        ),
                        "piece_weight_kg": (
                            float(
                                st.session_state.get(
                                    weight_key,
                                    saved_entry.get("piece_weight_kg", 0.0),
                                )
                                or 0.0
                            )
                            if operation_kind == "abkant"
                            else 0.0
                        ),
                        "bend_type": str(
                            st.session_state.get(
                                bend_key,
                                saved_entry.get("bend_type", ""),
                            )
                        ) if operation_kind == "abkant" else "",
                        "calculated_area_mm2": max(processed - fire, 0)
                        * float(item["unit_area_mm2"]),
                    })
        return entries

    # --------------------------------------------------------------
    # ADIM 1
    # --------------------------------------------------------------
    if step == 1:
        render_wizard_panel(
            "1 · Çalışan ve çalışma şekli",
            "Başlangıç çalışma şeklini seç. Gerçek toplam süre, üretim etaplarında girilen işlem sürelerinden otomatik hesaplanır.",
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
                "Başlangıç çalışma şekli",
                ["Tam zamanlı", "Yarı zamanlı", "Mesaili"],
                key=type_key,
                help="Toplam günlük süre 9 saati aşarsa veya tarih hafta sonuysa kayıt otomatik olarak Mesaili olur.",
            )

        reason = st.selectbox(
            "Yarı zamanlı çalışma nedeni",
            [""] + reasons,
            disabled=(work_type != "Yarı zamanlı"),
            key=reason_key,
        )

        selected_date_value = pd.to_datetime(production_date, errors="coerce")
        if not pd.isna(selected_date_value) and selected_date_value.weekday() >= 5:
            render_status_pill(
                "◷ Hafta sonu seçildi · Kayıt otomatik olarak Mesaili olacaktır",
                "purple",
            )
        elif work_type == "Mesaili":
            render_status_pill("◷ Mesaili kayıt · Yönetici bildirimi oluşur", "purple")
        else:
            st.caption(
                "İşlem sürelerinin günlük toplamı 9 saati aşarsa sistem çalışma şeklini otomatik Mesaili yapar."
            )

        if st.button("OC ve POS Seçimine Geç →", type="primary", use_container_width=True):
            errors = []
            if not operator_name:
                errors.append("Operatör / işçi seçmelisin.")
            if work_type == "Yarı zamanlı" and not reason:
                errors.append("Yarı zamanlı çalışma için neden seçmelisin.")
            if errors:
                for error in errors:
                    st.error(error)
            else:
                save_wizard_data(
                    production_date=production_date,
                    operator_name=operator_name,
                    declared_work_type=work_type,
                    work_type=work_type,
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
            "Çalışanın o gün gerçekten yaptığı etapları seç. Fazla üretim yapıldıysa tamamlanmış etap yeniden seçilebilir.",
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
                    normalized_operation = _normalize_header(operation)
                    operation_kind = _operation_kind(operation)

                    saved_entries_for_operation = [
                        saved_entry
                        for saved_entry in wizard_data.get("entries", [])
                        if int(saved_entry["item_id"]) == item_id
                        and str(saved_entry["operation_name"]) == operation
                    ]

                    st.markdown(f"**{operation}**")
                    overproduction_qty = max(completed_qty - requested_qty, 0)
                    if overproduction_qty > 0:
                        tone = "purple"
                        progress_text = (
                            f"{completed_qty}/{requested_qty} tamamlandı · "
                            f"{overproduction_qty} fazla üretildi"
                        )
                    else:
                        tone = (
                            "green"
                            if remaining_qty == 0
                            else "blue"
                            if completed_qty > 0
                            else "gray"
                        )
                        progress_text = (
                            f"{completed_qty}/{requested_qty} tamamlandı · "
                            f"{remaining_qty} kaldı"
                        )
                    render_status_pill(progress_text, tone)

                    if operation_kind == "laser":
                        lot_count_key = f"laserlotcount_{item_id}_{normalized_operation}_{version}"
                        if lot_count_key not in st.session_state:
                            st.session_state[lot_count_key] = max(
                                len(saved_entries_for_operation), 1
                            )
                        lot_count = max(int(st.session_state.get(lot_count_key, 1) or 1), 1)

                        st.markdown("### LAZER ÜRETİM GİRİŞİ")
                        st.caption(
                            "Her farklı malzeme veya kalınlık için ayrı bir üretim grubu oluştur. "
                            "Uzun-kısa kenar adetleriyle takım ve net m² otomatik hesaplanır; "
                            "kullanılan plaka adedi manuel girilir."
                        )

                        for lot_index in range(1, lot_count + 1):
                            saved_lot = (
                                saved_entries_for_operation[lot_index - 1]
                                if lot_index - 1 < len(saved_entries_for_operation)
                                else {}
                            )
                            suffix = f"laserlot_{lot_index}_{version}"
                            fire_key = f"opfire_{item_id}_{normalized_operation}_{suffix}"
                            hours_key = f"ophours_{item_id}_{normalized_operation}_{suffix}"
                            plate_key = f"laserplates_{item_id}_{normalized_operation}_{suffix}"
                            material_key = f"material_{item_id}_{normalized_operation}_{suffix}"
                            thickness_key = f"thickness_{item_id}_{normalized_operation}_{suffix}"
                            long_qty_key = f"laserlongqty_{item_id}_{normalized_operation}_{suffix}"
                            short_qty_key = f"lasershortqty_{item_id}_{normalized_operation}_{suffix}"
                            long_per_team_key = f"laserlongteam_{item_id}_{normalized_operation}_{suffix}"
                            short_per_team_key = f"lasershortteam_{item_id}_{normalized_operation}_{suffix}"

                            defaults_for_widgets = {
                                fire_key: int(saved_lot.get("fire_qty", 0)),
                                hours_key: float(saved_lot.get("operation_hours", 0.0)),
                                plate_key: int(saved_lot.get("laser_plate_qty", 0)),
                                material_key: str(saved_lot.get("material_type", "Alüminyum") or "Alüminyum"),
                                thickness_key: float(saved_lot.get("thickness_mm", 1.2) or 1.2),
                                long_qty_key: int(saved_lot.get("laser_long_edge_qty", 0)),
                                short_qty_key: int(saved_lot.get("laser_short_edge_qty", 0)),
                                long_per_team_key: int(saved_lot.get("laser_long_per_team", 2) or 2),
                                short_per_team_key: int(saved_lot.get("laser_short_per_team", 2) or 2),
                            }
                            for widget_key, widget_default in defaults_for_widgets.items():
                                if widget_key not in st.session_state:
                                    st.session_state[widget_key] = widget_default

                            st.markdown(f"##### Üretim grubu {lot_index}")
                            st.write(
                                f"**OC:** {selected_oc}  |  **POS:** {pos_name}  |  "
                                f"**Panel ölçüsü:** {_format_mm2(item['boy_mm'])} × {_format_mm2(item['en_mm'])} mm"
                            )

                            material_col, thickness_col, plate_col, hours_col, fire_col = st.columns(5)
                            with material_col:
                                st.selectbox(
                                    "Malzeme",
                                    ["Çelik", "Alüminyum", "Paslanmaz"],
                                    key=material_key,
                                )
                            with thickness_col:
                                st.number_input(
                                    "Kalınlık (mm)",
                                    min_value=0.1,
                                    max_value=100.0,
                                    step=0.1,
                                    key=thickness_key,
                                )
                            with plate_col:
                                st.number_input(
                                    "Plaka adedi",
                                    min_value=0,
                                    step=1,
                                    key=plate_key,
                                    help="Lazerde kullanılan plaka sayısını manuel girin.",
                                )
                            with hours_col:
                                st.number_input(
                                    "Süre (saat)",
                                    min_value=0.0,
                                    max_value=24.0,
                                    step=0.25,
                                    key=hours_key,
                                )
                            with fire_col:
                                st.number_input(
                                    "Fire takım",
                                    min_value=0,
                                    step=1,
                                    key=fire_key,
                                    help="Tamamlanan takımlardan fire olan takım sayısı.",
                                )

                            edge1, edge2 = st.columns(2)
                            with edge1:
                                st.number_input(
                                    "Uzun kenar üretimi (adet)",
                                    min_value=0,
                                    step=1,
                                    key=long_qty_key,
                                )
                            with edge2:
                                st.number_input(
                                    "Kısa kenar üretimi (adet)",
                                    min_value=0,
                                    step=1,
                                    key=short_qty_key,
                                )

                            st.markdown("**Bir takım için gereken kenarlar**")
                            team1, team2 = st.columns(2)
                            with team1:
                                st.number_input(
                                    "Uzun kenar / takım",
                                    min_value=1,
                                    step=1,
                                    key=long_per_team_key,
                                )
                            with team2:
                                st.number_input(
                                    "Kısa kenar / takım",
                                    min_value=1,
                                    step=1,
                                    key=short_per_team_key,
                                )

                            long_qty = int(st.session_state.get(long_qty_key, 0) or 0)
                            short_qty = int(st.session_state.get(short_qty_key, 0) or 0)
                            long_per_team = max(int(st.session_state.get(long_per_team_key, 2) or 2), 1)
                            short_per_team = max(int(st.session_state.get(short_per_team_key, 2) or 2), 1)
                            completed_teams = min(
                                long_qty // long_per_team,
                                short_qty // short_per_team,
                            )
                            fire_teams = int(st.session_state.get(fire_key, 0) or 0)
                            good_teams = max(completed_teams - fire_teams, 0)
                            long_excess = long_qty - completed_teams * long_per_team
                            short_excess = short_qty - completed_teams * short_per_team
                            unit_area_m2 = float(item["unit_area_mm2"]) / 1_000_000
                            net_area_m2 = good_teams * unit_area_m2
                            current_hours = float(st.session_state.get(hours_key, 0.0) or 0.0)
                            m2_hour = net_area_m2 / current_hours if current_hours > 0 else 0.0

                            manual_plate_qty = int(
                                st.session_state.get(plate_key, 0) or 0
                            )
                            result1, result2, result3, result4 = st.columns(4)
                            result1.metric("Plaka adedi", f"{manual_plate_qty} plaka")
                            result2.metric("Tamamlanan takım", f"{completed_teams} takım")
                            result3.metric("Sağlam takım", f"{good_teams} takım")
                            result4.metric(
                                "Üretilen net panel alanı",
                                f"{net_area_m2:.2f} m²",
                            )

                            st.info(
                                f"**Eksik / fazla parça:** Uzun kenar {long_excess} adet fazla · "
                                f"Kısa kenar {short_excess} adet fazla"
                            )
                            st.markdown(
                                f"""**1 panel alanı:**  
{_format_mm2(item['boy_mm'])} × {_format_mm2(item['en_mm'])} ÷ 1.000.000 = **{unit_area_m2:.2f} m²**

**Tam takım:** {completed_teams}

**Üretilen net panel alanı:**  
{good_teams} × {unit_area_m2:.2f} = **{net_area_m2:.2f} m²**"""
                            )
                            st.caption(f"Alan verimi: {m2_hour:.2f} m²/saat")
                            if lot_index < lot_count:
                                st.markdown("---")

                        add_lot_col, remove_lot_col = st.columns([1.6, 1.0])
                        with add_lot_col:
                            if st.button(
                                "＋ Yeni Malzeme / Kalınlık Grubu Ekle",
                                key=f"add_laser_lot_{item_id}_{normalized_operation}_{version}",
                                use_container_width=True,
                                type="primary",
                            ):
                                st.session_state[lot_count_key] = lot_count + 1
                                st.rerun()
                        with remove_lot_col:
                            if st.button(
                                "− Son Grubu Sil",
                                key=f"remove_laser_lot_{item_id}_{normalized_operation}_{version}",
                                use_container_width=True,
                                disabled=(lot_count <= 1),
                            ):
                                removed_suffix = f"laserlot_{lot_count}_{version}"
                                for removable_key in (
                                    f"opfire_{item_id}_{normalized_operation}_{removed_suffix}",
                                    f"ophours_{item_id}_{normalized_operation}_{removed_suffix}",
                                    f"material_{item_id}_{normalized_operation}_{removed_suffix}",
                                    f"thickness_{item_id}_{normalized_operation}_{removed_suffix}",
                                    f"laserplates_{item_id}_{normalized_operation}_{removed_suffix}",
                                    f"laserlongqty_{item_id}_{normalized_operation}_{removed_suffix}",
                                    f"lasershortqty_{item_id}_{normalized_operation}_{removed_suffix}",
                                    f"laserlongteam_{item_id}_{normalized_operation}_{removed_suffix}",
                                    f"lasershortteam_{item_id}_{normalized_operation}_{removed_suffix}",
                                ):
                                    st.session_state.pop(removable_key, None)
                                st.session_state[lot_count_key] = lot_count - 1
                                st.rerun()

                    elif operation_kind == "laser_cut_out":
                        processed_key = (
                            f"processed_{item_id}_{normalized_operation}_{version}"
                        )
                        saved_entry = (
                            saved_entries_for_operation[0]
                            if saved_entries_for_operation
                            else {}
                        )
                        if processed_key not in st.session_state:
                            st.session_state[processed_key] = int(
                                saved_entry.get("processed_qty", 0) or 0
                            )

                        cutout_col1, cutout_col2 = st.columns([1.4, 1.0])
                        with cutout_col1:
                            st.number_input(
                                "Üretilen adet",
                                min_value=0,
                                step=1,
                                key=processed_key,
                            )
                        with cutout_col2:
                            st.write("")
                            st.button(
                                "Kalanı yaz",
                                disabled=(remaining_qty == 0),
                                use_container_width=True,
                                key=(
                                    f"fill_cutout_{item_id}_"
                                    f"{normalized_operation}_{version}"
                                ),
                                on_click=set_session_value,
                                args=(processed_key, remaining_qty),
                            )

                        cutout_qty = int(
                            st.session_state.get(processed_key, 0) or 0
                        )
                        cutout_area_m2 = (
                            cutout_qty * float(item["unit_area_mm2"]) / 1_000_000
                        )
                        st.info(
                            f"Laser Cut Out: **{cutout_qty} adet · {cutout_area_m2:.1f} m²**"
                        )
                        st.caption(
                            "Bu etapta yalnızca adet girilir. Plaka, malzeme ve kalınlık "
                            "alanı yoktur; m² POS ölçüsünden otomatik hesaplanır."
                        )

                    else:
                        processed_key = f"processed_{item_id}_{normalized_operation}_{version}"
                        fire_key = f"opfire_{item_id}_{normalized_operation}_{version}"
                        hours_key = f"ophours_{item_id}_{normalized_operation}_{version}"
                        work_mode_key = f"abkantmode_{item_id}_{normalized_operation}_{version}"
                        coworker_key = f"abkantcoworker_{item_id}_{normalized_operation}_{version}"
                        abkant_team_qty_key = (
                            f"abkantteamqty_{item_id}_{normalized_operation}_{version}"
                        )
                        abkant_teams_per_piece_key = (
                            f"abkantteamsperpiece_{item_id}_{normalized_operation}_{version}"
                        )
                        abkant_long_bend_qty_key = (
                            f"abkantlongbendqty_{item_id}_{normalized_operation}_{version}"
                        )
                        abkant_short_bend_qty_key = (
                            f"abkantshortbendqty_{item_id}_{normalized_operation}_{version}"
                        )
                        abkant_long_single_bend_qty_key = (
                            f"abkantlongsinglebendqty_{item_id}_{normalized_operation}_{version}"
                        )
                        abkant_long_double_bend_qty_key = (
                            f"abkantlongdoublebendqty_{item_id}_{normalized_operation}_{version}"
                        )
                        abkant_short_single_bend_qty_key = (
                            f"abkantshortsinglebendqty_{item_id}_{normalized_operation}_{version}"
                        )
                        abkant_short_double_bend_qty_key = (
                            f"abkantshortdoublebendqty_{item_id}_{normalized_operation}_{version}"
                        )
                        abkant_manual_override_key = (
                            f"abkantmanualoverride_{item_id}_{normalized_operation}_{version}"
                        )
                        participants_key = f"participants_{item_id}_{normalized_operation}_{version}"
                        weight_key = f"pieceweight_{item_id}_{normalized_operation}_{version}"
                        bend_key = f"bendtype_{item_id}_{normalized_operation}_{version}"
                        saved_entry = (
                            saved_entries_for_operation[0]
                            if saved_entries_for_operation
                            else {}
                        )
                        defaults_for_widgets = {
                            processed_key: int(saved_entry.get("processed_qty", 0)),
                            abkant_teams_per_piece_key: max(
                                int(saved_entry.get("abkant_teams_per_piece", 1) or 1),
                                1,
                            ),
                            abkant_team_qty_key: int(
                                saved_entry.get(
                                    "abkant_team_qty",
                                    int(saved_entry.get("processed_qty", 0) or 0)
                                    * max(
                                        int(
                                            saved_entry.get(
                                                "abkant_teams_per_piece",
                                                1,
                                            )
                                            or 1
                                        ),
                                        1,
                                    ),
                                )
                                or 0
                            ),
                            abkant_long_bend_qty_key: int(
                                saved_entry.get("abkant_long_bend_qty", 0) or 0
                            ),
                            abkant_short_bend_qty_key: int(
                                saved_entry.get("abkant_short_bend_qty", 0) or 0
                            ),
                            abkant_long_single_bend_qty_key: int(
                                saved_entry.get("abkant_long_single_bend_qty", 0) or 0
                            ),
                            abkant_long_double_bend_qty_key: int(
                                saved_entry.get("abkant_long_double_bend_qty", 0) or 0
                            ),
                            abkant_short_single_bend_qty_key: int(
                                saved_entry.get("abkant_short_single_bend_qty", 0) or 0
                            ),
                            abkant_short_double_bend_qty_key: int(
                                saved_entry.get("abkant_short_double_bend_qty", 0) or 0
                            ),
                            abkant_manual_override_key: bool(
                                saved_entry.get("abkant_manual_override", False)
                            ),
                            fire_key: int(saved_entry.get("fire_qty", 0)),
                            hours_key: float(saved_entry.get("operation_hours", 0.0)),
                            work_mode_key: str(saved_entry.get("abkant_work_mode", "Tek çalıştı") or "Tek çalıştı"),
                            coworker_key: str(saved_entry.get("abkant_coworker", "") or ""),
                            participants_key: _split_participants(
                                saved_entry.get("participants_text", "")
                            ),
                            weight_key: float(saved_entry.get("piece_weight_kg", 0.0) or 0.0),
                            bend_key: str(saved_entry.get("bend_type", "Tek büküm") or "Tek büküm"),
                        }
                        for widget_key, widget_default in defaults_for_widgets.items():
                            if widget_key not in st.session_state:
                                st.session_state[widget_key] = widget_default

                        if operation_kind == "abkant":
                            st.markdown("**Kenar ve büküm türü dağılımı**")
                            bend1, bend2, bend3, bend4 = st.columns(4)
                            with bend1:
                                st.number_input(
                                    "Uzun kenar · Tek büküm", min_value=0, step=1,
                                    key=abkant_long_single_bend_qty_key,
                                )
                                st.button(
                                    "Kalanı yaz",
                                    disabled=(remaining_qty == 0),
                                    use_container_width=True,
                                    key=(
                                        f"fill_abkant_long_single_{item_id}_"
                                        f"{normalized_operation}_{version}"
                                    ),
                                    on_click=set_session_value,
                                    args=(abkant_long_single_bend_qty_key, remaining_qty),
                                )
                            with bend2:
                                st.number_input(
                                    "Uzun kenar · Çift büküm", min_value=0, step=1,
                                    key=abkant_long_double_bend_qty_key,
                                )
                                st.button(
                                    "Kalanı yaz",
                                    disabled=(remaining_qty == 0),
                                    use_container_width=True,
                                    key=(
                                        f"fill_abkant_long_double_{item_id}_"
                                        f"{normalized_operation}_{version}"
                                    ),
                                    on_click=set_session_value,
                                    args=(abkant_long_double_bend_qty_key, remaining_qty),
                                )
                            with bend3:
                                st.number_input(
                                    "Kısa kenar · Tek büküm", min_value=0, step=1,
                                    key=abkant_short_single_bend_qty_key,
                                )
                                st.button(
                                    "Kalanı yaz",
                                    disabled=(remaining_qty == 0),
                                    use_container_width=True,
                                    key=(
                                        f"fill_abkant_short_single_{item_id}_"
                                        f"{normalized_operation}_{version}"
                                    ),
                                    on_click=set_session_value,
                                    args=(abkant_short_single_bend_qty_key, remaining_qty),
                                )
                            with bend4:
                                st.number_input(
                                    "Kısa kenar · Çift büküm", min_value=0, step=1,
                                    key=abkant_short_double_bend_qty_key,
                                )
                                st.button(
                                    "Kalanı yaz",
                                    disabled=(remaining_qty == 0),
                                    use_container_width=True,
                                    key=(
                                        f"fill_abkant_short_double_{item_id}_"
                                        f"{normalized_operation}_{version}"
                                    ),
                                    on_click=set_session_value,
                                    args=(abkant_short_double_bend_qty_key, remaining_qty),
                                )
                            st.caption(
                                "Bugün hangi kenar üzerinde çalışıldıysa yalnızca o alanı gir. "
                                "Diğer kenarlar 0 kalabilir. Sistem geçmiş kayıtlarla birleştirerek "
                                "tamamlanan ürün adedini otomatik hesaplar."
                            )

                            qty1, qty2 = st.columns(2)
                            with qty1:
                                st.number_input(
                                    "Fire (adet)",
                                    min_value=0,
                                    step=1,
                                    key=fire_key,
                                    help="Yalnızca bu kayıtta tamamlanan ürünlerden fire çıkan varsa girin.",
                                )
                            with qty2:
                                st.number_input(
                                    "Süre (saat)",
                                    min_value=0.0,
                                    max_value=24.0,
                                    step=0.25,
                                    key=hours_key,
                                )

                            current_long_qty = (
                                int(st.session_state.get(abkant_long_single_bend_qty_key, 0) or 0)
                                + int(st.session_state.get(abkant_long_double_bend_qty_key, 0) or 0)
                            )
                            current_short_qty = (
                                int(st.session_state.get(abkant_short_single_bend_qty_key, 0) or 0)
                                + int(st.session_state.get(abkant_short_double_bend_qty_key, 0) or 0)
                            )
                            today_possible_complete = min(current_long_qty, current_short_qty)
                            st.info(
                                f"Bugünkü giriş: **{current_long_qty} uzun kenar · "
                                f"{current_short_qty} kısa kenar**. "
                                f"Geçmiş kayıtlar hesaba katılmadan bugün kendi içinde "
                                f"en fazla **{today_possible_complete} adet** tamamlanabilir."
                            )

                            abkant1, abkant2, abkant3, abkant4 = st.columns(4)
                            with abkant1:
                                st.selectbox(
                                    "Çalışma şekli",
                                    ["Tek çalıştı", "Biriyle beraber çalıştı"],
                                    key=work_mode_key,
                                )
                            with abkant2:
                                st.text_input(
                                    "Büküm türü",
                                    value="Kenar bazında yukarıdan girildi",
                                    disabled=True,
                                    key=f"bend_summary_{item_id}_{normalized_operation}_{version}",
                                )
                            with abkant3:
                                coworker_options = [""] + [
                                    name for name in WORKER_NAMES
                                    if name and name != operator_name
                                ]
                                st.selectbox(
                                    "Beraber çalıştığı kişi",
                                    coworker_options,
                                    disabled=(
                                        st.session_state.get(work_mode_key)
                                        != "Biriyle beraber çalıştı"
                                    ),
                                    format_func=lambda value: "Kişi seçin" if value == "" else value,
                                    key=coworker_key,
                                )
                            with abkant4:
                                st.number_input(
                                    "Tek parça ağırlığı (kg)",
                                    min_value=0.0,
                                    max_value=500.0,
                                    step=0.5,
                                    key=weight_key,
                                    help="Gerçek parça ağırlığını girin. Yönetici iki kişi çalışma kontrolünü bu değerle yapar.",
                                )

                            long_edge_mm = max(float(item["boy_mm"]), float(item["en_mm"]))
                            piece_weight = float(st.session_state.get(weight_key, 0.0) or 0.0)
                            if long_edge_mm > 1500 and piece_weight > 15:
                                if st.session_state.get(work_mode_key) == "Biriyle beraber çalıştı":
                                    st.success(
                                        f"İki kişi kuralı sağlanıyor: uzun kenar {long_edge_mm:.0f} mm, "
                                        f"ağırlık {piece_weight:.1f} kg."
                                    )
                                else:
                                    st.error(
                                        f"İKİ KİŞİ ÇALIŞMALI: uzun kenar {long_edge_mm:.0f} mm ve "
                                        f"ağırlık {piece_weight:.1f} kg."
                                    )
                        else:
                            row1, row2, row3, row4 = st.columns(
                                [1.2, 1.0, 1.0, 1.0]
                            )
                            with row1:
                                st.number_input(
                                    "İşlem yapılan",
                                    min_value=0,
                                    step=1,
                                    key=processed_key,
                                )
                            with row2:
                                st.number_input(
                                    "Fire",
                                    min_value=0,
                                    step=1,
                                    key=fire_key,
                                )
                            with row3:
                                st.number_input(
                                    "Süre (saat)",
                                    min_value=0.0,
                                    max_value=24.0,
                                    step=0.25,
                                    key=hours_key,
                                )
                            with row4:
                                st.write("")
                                st.button(
                                    "Kalanı yaz",
                                    disabled=(remaining_qty == 0),
                                    use_container_width=True,
                                    key=(
                                        f"fill_{item_id}_"
                                        f"{normalized_operation}_{version}"
                                    ),
                                    on_click=set_session_value,
                                    args=(processed_key, remaining_qty),
                                )

                        if operation_kind in {"boya", "paketleme", "sevkiyat"}:
                            team_options = [
                                name for name in WORKER_NAMES
                                if name and name != operator_name
                            ]
                            operation_team_label = {
                                "boya": "Boya çalışma ekibi",
                                "paketleme": "Paketleme çalışma ekibi",
                                "sevkiyat": "Sevkiyat çalışma ekibi",
                            }[operation_kind]
                            st.multiselect(
                                operation_team_label,
                                team_options,
                                key=participants_key,
                                help=(
                                    "Ana çalışan dışında bu işe katılan kişileri seçin. "
                                    "Seçilen kişiler çalışan bazlı özetlerde de görünür."
                                ),
                            )

                        if operation_kind == "abkant":
                            current_team_qty = int(
                                st.session_state.get(abkant_team_qty_key, 0) or 0
                            )
                            current_teams_per_piece = max(
                                int(
                                    st.session_state.get(
                                        abkant_teams_per_piece_key,
                                        1,
                                    )
                                    or 1
                                ),
                                1,
                            )
                            calculated_processed = (
                                current_team_qty // current_teams_per_piece
                            )
                            current_processed = (
                                int(st.session_state.get(processed_key, 0) or 0)
                                if bool(
                                    st.session_state.get(
                                        abkant_manual_override_key, False
                                    )
                                )
                                else calculated_processed
                            )
                        else:
                            current_processed = int(
                                st.session_state.get(processed_key, 0) or 0
                            )
                        current_fire = int(
                            st.session_state.get(fire_key, 0) or 0
                        )
                        current_good = max(current_processed - current_fire, 0)
                        current_area_mm2 = current_good * float(item["unit_area_mm2"])
                        current_area_m2 = current_area_mm2 / 1_000_000
                        current_hours = float(
                            st.session_state.get(hours_key, 0.0) or 0.0
                        )
                        current_m2_hour = (
                            current_area_m2 / current_hours
                            if current_hours > 0
                            else 0.0
                        )
                        st.info(
                            f"Bu etapta işçinin yaptığı sağlam alan: **{current_area_m2:.1f} m²** "
                            f"({current_good} sağlam adet)"
                        )
                        st.caption(
                            f"Birim parça: {_format_m2(item['unit_area_mm2'])} m² · "
                            f"Etap verimi: {current_m2_hour:.1f} m²/saat"
                        )
                    st.divider()
        entries = build_entries(selected_items, plans, selected_map)
        entered_work_hours = _entries_total_hours(entries)
        performance = calculate_entries_performance(entries, entered_work_hours)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("İşlem yapılan", f"{performance['total_processed']} adet")
        m2.metric("Sağlam ilerleyen", f"{performance['total_good']} adet")
        m3.metric("Fire", f"{performance['total_fire']} adet")
        m4.metric("Girilen toplam süre", f"{entered_work_hours:.2f} saat")

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
                if entered_work_hours > 24.0:
                    errors.append("Girilen işlem sürelerinin toplamı bir günde 24 saati aşamaz.")
                for entry in entries:
                    if entry["fire_qty"] > entry["processed_qty"]:
                        errors.append(
                            f"{entry['pos']} / {entry['operation_name']}: Fire işlenen adetten büyük olamaz."
                        )
                    operation_kind = _operation_kind(entry["operation_name"])
                    if (
                        operation_kind != "laser_cut_out"
                        and float(entry.get("operation_hours", 0) or 0) <= 0
                    ):
                        errors.append(
                            f"{entry['pos']} / {entry['operation_name']}: İşlem süresini girmelisin."
                        )
                    if operation_kind == "laser":
                        if int(entry.get("laser_plate_qty", 0) or 0) <= 0:
                            errors.append(
                                f"{entry['pos']} / {entry['operation_name']}: Plaka adedini girmelisin."
                            )
                        if not str(entry.get("material_type", "")).strip():
                            errors.append(
                                f"{entry['pos']} / {entry['operation_name']}: Malzeme seçmelisin."
                            )
                        if float(entry.get("thickness_mm", 0) or 0) <= 0:
                            errors.append(
                                f"{entry['pos']} / {entry['operation_name']}: Malzeme kalınlığını girmelisin."
                            )
                    if operation_kind == "abkant":
                        abkant_edge_total = sum(
                            max(int(entry.get(field, 0) or 0), 0)
                            for field in (
                                "abkant_long_single_bend_qty",
                                "abkant_long_double_bend_qty",
                                "abkant_short_single_bend_qty",
                                "abkant_short_double_bend_qty",
                            )
                        )
                        if abkant_edge_total <= 0:
                            errors.append(
                                f"{entry['pos']} / {entry['operation_name']}: "
                                "Bugün yapılan uzun veya kısa kenar büküm adedinden en az birini girmelisin."
                            )
                        if (
                            entry.get("abkant_work_mode") == "Biriyle beraber çalıştı"
                            and not str(entry.get("abkant_coworker", "")).strip()
                        ):
                            errors.append(
                                f"{entry['pos']} / {entry['operation_name']}: Beraber çalıştığı kişiyi seçmelisin."
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

    declared_work_type = str(
        wizard_data.get(
            "declared_work_type",
            wizard_data.get(
                "work_type",
                st.session_state.get(type_key, "Tam zamanlı"),
            ),
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
    operator_name = str(
        wizard_data.get(
            "operator_name",
            operator_name,
        )
    )
    work_hours = _entries_total_hours(entries)
    # Günlük saat sınırı kritik bir kontroldür; önbellek kullanmadan
    # doğrudan veritabanındaki mevcut kayıtları okur.
    existing_daily_hours = get_registered_daily_hours(
        operator_name,
        production_date,
    )
    cumulative_hours = round(existing_daily_hours + work_hours, 2)
    work_type = determine_work_type(
        production_date,
        declared_work_type,
        work_hours,
        existing_daily_hours,
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
    if work_type == "Mesaili":
        reason_parts = []
        parsed_production_date = pd.to_datetime(production_date, errors="coerce")
        if not pd.isna(parsed_production_date) and parsed_production_date.weekday() >= 5:
            reason_parts.append("hafta sonu çalışması")
        if cumulative_hours > 9.0:
            reason_parts.append(f"günlük toplam {cumulative_hours:.2f} saat")
        if declared_work_type == "Mesaili":
            reason_parts.append("başlangıçta mesaili seçildi")
        st.warning(
            "Çalışma şekli Mesaili olarak kaydedilecek"
            + (" · " + ", ".join(reason_parts) if reason_parts else "")
        )

    st.caption(
        f"Günlük süre hesabı: Önceden kayıtlı {existing_daily_hours:.1f} saat "
        f"+ bu kayıt {work_hours:.1f} saat = toplam {cumulative_hours:.1f} saat."
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
            "Fazla Üretim": max(entry["good_qty"] - entry["remaining_before"], 0),
            "Süre (saat)": round(float(entry.get("operation_hours", 0)), 2),
            "Malzeme": entry.get("material_type", ""),
            "Kalınlık (mm)": entry.get("thickness_mm", 0) or "",
            "Lazer Grubu": entry.get("laser_lot_no", 0) or "",
            "Lazer Plaka": entry.get("laser_plate_qty", 0) or "",
            "Abkant Takım": entry.get("abkant_team_qty", 0) or "",
            "Takım / Adet": entry.get("abkant_teams_per_piece", 0) or "",
            "Artan Takım": entry.get("abkant_team_excess", 0) or "",
            "Uzun Tek Büküm": entry.get("abkant_long_single_bend_qty", 0) or "",
            "Uzun Çift Büküm": entry.get("abkant_long_double_bend_qty", 0) or "",
            "Kısa Tek Büküm": entry.get("abkant_short_single_bend_qty", 0) or "",
            "Kısa Çift Büküm": entry.get("abkant_short_double_bend_qty", 0) or "",
            "Adet Manuel Düzeltildi": (
                "Evet" if entry.get("abkant_manual_override") else ""
            ),
            "Abkant Çalışma": entry.get("abkant_work_mode", ""),
            "Beraber Çalıştığı": entry.get("abkant_coworker", ""),
            "Çalışma Ekibi": entry.get("participants_text", ""),
            "Parça Ağırlığı (kg)": entry.get("piece_weight_kg", 0) or "",
            "Büküm Türü": entry.get("bend_type", ""),
            "Sağlam Alan (m²)": round(
                entry["good_qty"] * entry["unit_area_mm2"] / 1_000_000,
                1,
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

    heavy_single_abkant = [
        entry for entry in entries
        if _operation_kind(entry.get("operation_name", "")) == "abkant"
        and max(float(entry.get("boy_mm", 0)), float(entry.get("en_mm", 0))) > 1500
        and float(entry.get("piece_weight_kg", 0) or 0) > 15
        and entry.get("abkant_work_mode") != "Biriyle beraber çalıştı"
    ]
    for entry in heavy_single_abkant:
        st.error(
            f"{entry['pos']} için iki kişi çalışma uyarısı: "
            f"uzun kenar {max(float(entry.get('boy_mm', 0)), float(entry.get('en_mm', 0))):.0f} mm, "
            f"ağırlık {float(entry.get('piece_weight_kg', 0)):.1f} kg."
        )

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
            if cumulative_hours > 24.0001:
                final_errors.append(
                    "Bu çalışan için günlük toplam çalışma süresi 24 saati aşamaz. "
                    f"Önceden kayıtlı: {existing_daily_hours:.1f} saat, "
                    f"bu kayıt: {work_hours:.1f} saat, "
                    f"toplam: {cumulative_hours:.1f} saat."
                )
            if work_type == "Yarı zamanlı" and not reason:
                final_errors.append("Yarı zamanlı çalışma için neden seçmelisin.")
            for entry in entries:
                if entry["fire_qty"] > entry["processed_qty"]:
                    final_errors.append(
                        f"{entry['pos']} / {entry['operation_name']}: Fire işlenen adetten büyük olamaz."
                    )
                operation_kind = _operation_kind(entry["operation_name"])
                if (
                    operation_kind != "laser_cut_out"
                    and float(entry.get("operation_hours", 0) or 0) <= 0
                ):
                    final_errors.append(
                        f"{entry['pos']} / {entry['operation_name']}: İşlem süresini girmelisin."
                    )
                if operation_kind == "laser":
                    if int(entry.get("laser_plate_qty", 0) or 0) <= 0:
                        final_errors.append(
                            f"{entry['pos']} / {entry['operation_name']}: Plaka adedini girmelisin."
                        )
                    if not str(entry.get("material_type", "")).strip():
                        final_errors.append(
                            f"{entry['pos']} / {entry['operation_name']}: Malzeme seçmelisin."
                        )
                    if float(entry.get("thickness_mm", 0) or 0) <= 0:
                        final_errors.append(
                            f"{entry['pos']} / {entry['operation_name']}: Malzeme kalınlığını girmelisin."
                        )
                if operation_kind == "abkant":
                    abkant_edge_total = sum(
                        max(int(entry.get(field, 0) or 0), 0)
                        for field in (
                            "abkant_long_single_bend_qty",
                            "abkant_long_double_bend_qty",
                            "abkant_short_single_bend_qty",
                            "abkant_short_double_bend_qty",
                        )
                    )
                    if abkant_edge_total <= 0:
                        final_errors.append(
                            f"{entry['pos']} / {entry['operation_name']}: "
                            "Bugün yapılan uzun veya kısa kenar büküm adedinden en az birini girmelisin."
                        )
                    if (
                        entry.get("abkant_work_mode") == "Biriyle beraber çalıştı"
                        and not str(entry.get("abkant_coworker", "")).strip()
                    ):
                        final_errors.append(
                            f"{entry['pos']} / {entry['operation_name']}: Beraber çalıştığı kişiyi seçmelisin."
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

    total_requested_qty = int(filtered["requested_qty"].sum())
    total_produced_qty = int(filtered["produced_qty"].sum())
    total_requested_m2 = float(filtered["requested_area_mm2"].sum()) / 1_000_000
    total_produced_m2 = float(filtered["produced_area_mm2"].sum()) / 1_000_000
    qty_ratio = (total_produced_qty / total_requested_qty * 100) if total_requested_qty else 0.0
    m2_ratio = (total_produced_m2 / total_requested_m2 * 100) if total_requested_m2 else 0.0

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("İstenen", total_requested_qty)
    k2.metric("Tam biten", total_produced_qty)
    k3.metric("Fazla üretim", int(filtered["overproduction_qty"].sum()))
    k4.metric("Üretimde olan", int(filtered["in_production_qty"].sum()))
    k5.metric("Operasyon firesi", int(filtered["operation_fire_qty"].sum()))

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Üretilen adet / toplam adet", f"%{qty_ratio:.1f}", f"{total_produced_qty} / {total_requested_qty}")
    r2.metric("Üretilen m² / toplam m²", f"%{m2_ratio:.1f}", f"{total_produced_m2:.1f} / {total_requested_m2:.1f} m²")
    r3.metric("Toplam proje alanı", f"{total_requested_m2:.1f} m²")
    r4.metric("Üretilen alan", f"{total_produced_m2:.1f} m²")

    summary = filtered[[
        "oc_no", "pos", "combination_name", "requested_qty", "produced_qty",
        "overproduction_qty", "in_production_qty", "remaining_qty",
        "operation_fire_qty", "completion_pct",
    ]].rename(columns={
        "oc_no": "OC",
        "pos": "POS",
        "combination_name": "Kombinasyon",
        "requested_qty": "İstenen",
        "produced_qty": "Tam Biten",
        "overproduction_qty": "Fazla Üretim",
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
            "overproduction_qty", "remaining_qty", "processed_qty", "fire_qty",
            "completion_pct",
        ]].rename(columns={
            "operation_order": "Sıra",
            "operation_name": "İşlem",
            "requested_qty": "İstenen",
            "completed_qty": "Sağlam İlerleyen",
            "overproduction_qty": "Fazla Üretim",
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
            "Geçmiş m²/saat",
            f"{history_area / total_history_hours / 1_000_000:.1f}"
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
                            preview = parsed.copy()
                            preview["unit_area_m2"] = (
                                preview["unit_area_mm2"] / 1_000_000
                            ).round(1)
                            preview = preview.drop(
                                columns=["unit_area_mm2"], errors="ignore"
                            ).rename(columns={
                                "project_name": "Proje",
                                "oc_no": "OC",
                                "pos": "POS",
                                "requested_qty": "İstenen Adet",
                                "boy_mm": "Boy (mm)",
                                "en_mm": "En (mm)",
                                "unit_area_m2": "Birim Alan (m²)",
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
    total_overproduction = int(items["overproduction_qty"].sum())
    requested_area = float(items["requested_area_mm2"].sum())
    produced_area = float(items["produced_area_mm2"].sum())
    remaining_area = float(items["remaining_area_mm2"].sum())
    overproduction_area = float(items["overproduction_area_mm2"].sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("İstenen toplam adet", total_requested)
    m2.metric("Üretilen toplam adet", total_produced)
    m3.metric("Kalan toplam adet", total_remaining)
    m4.metric("Fazla üretim", total_overproduction)

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("İstenen alan", f"{_format_m2(requested_area)} m²")
    a2.metric("Üretilen alan", f"{_format_m2(produced_area)} m²")
    a3.metric("Kalan alan", f"{_format_m2(remaining_area)} m²")
    a4.metric("Fazla üretim alanı", f"{_format_m2(overproduction_area)} m²")

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
            help=(
                "Tamamlanmış bir POS için fazla üretim gireceksen bu seçeneği kapat."
            ),
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
            "Birim Alan (m²)": (
                editable["unit_area_mm2"] / 1_000_000
            ).round(1),
            "Önceden Üretilen": editable["produced_qty"].astype(int),
            "Kalan": editable["remaining_qty"].astype(int),
            "Mevcut Fazla": editable["overproduction_qty"].astype(int),
            "Bugün Üretilen": 0,
        })

        edited = st.data_editor(
            editor,
            use_container_width=True,
            hide_index=True,
            disabled=[
                "OC", "POS", "İstenen Adet", "Boy (mm)", "En (mm)",
                "Birim Alan (m²)", "Önceden Üretilen", "Kalan", "Mevcut Fazla",
            ],
            column_config={
                "item_id": None,
                "Bugün Üretilen": st.column_config.NumberColumn(
                    "Bugün Üretilen",
                    min_value=0,
                    step=1,
                    required=True,
                ),
                "Birim Alan (m²)": st.column_config.NumberColumn(format="%.1f"),
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
            (entered * edited["Birim Alan (m²)"].astype(float)).sum()
        )
        projected_overproduction = (
            edited["Önceden Üretilen"].astype(int)
            + entered
            - edited["İstenen Adet"].astype(int)
        ).clip(lower=0)

        e1, e2 = st.columns(2)
        e1.metric("Bugün girilen adet", int(entered.sum()))
        e2.metric(
            "Bugün üretilen alan",
            f"{entered_area:.1f} m²",
        )

        excess_rows = edited[projected_overproduction > 0].copy()
        if not excess_rows.empty:
            excess_rows["Fazla"] = projected_overproduction[
                projected_overproduction > 0
            ].astype(int)
            st.warning(
                "Fazla üretim olarak kaydedilecek: "
                + ", ".join(
                    f"{row['POS']} (+{int(row['Fazla'])})"
                    for _, row in excess_rows.iterrows()
                )
            )

        if st.button(
            "Üretim Miktarlarını Kaydet",
            type="primary",
            key=f"save_output_{selected_oc}",
        ):
            errors = []
            if not operator_name:
                errors.append("Operatör / işçi seçmelisin.")

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

    conn = get_db_connection()
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

    conn = get_db_connection()
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


def update_operation_batch_record(
    batch_id: int,
    batch_data: dict,
    edited_rows: pd.DataFrame,
) -> dict:
    """Ana kayıt ile işlem satırlarını ekleme, çıkarma ve düzenleme yoluyla günceller."""
    if edited_rows.empty:
        raise ValueError("Kayıtta en az bir işlem satırı bulunmalıdır.")

    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT tarih, operator_ismi FROM operation_batches WHERE id = ?",
            (int(batch_id),),
        )
        previous_batch = cur.fetchone()
        if previous_batch is None:
            raise ValueError("Güncellenecek operasyon kaydı bulunamadı.")

        existing_rows = cur.execute(
            "SELECT id, item_id FROM operation_work_logs WHERE batch_id = ?",
            (int(batch_id),),
        ).fetchall()
        allowed_ids = {int(row[0]) for row in existing_rows}
        affected_item_ids = {int(row[1]) for row in existing_rows}

        valid_types = {"Tam zamanlı", "Yarı zamanlı", "Mesaili"}
        work_type = str(batch_data.get("calisma_tipi", "Tam zamanlı"))
        if work_type not in valid_types:
            work_type = "Tam zamanlı"

        work_date = str(batch_data.get("tarih", ""))
        operator_name = str(batch_data.get("operator_ismi", "")).strip()
        work_hours = max(float(batch_data.get("calisma_saati", 0) or 0), 0.0)
        if not operator_name:
            raise ValueError("Çalışan adı boş bırakılamaz.")
        if work_hours <= 0:
            raise ValueError("Çalışma saati 0'dan büyük olmalıdır.")

        parsed_date = pd.to_datetime(work_date, errors="coerce")
        other_operation_hours = float(cur.execute(
            """
            SELECT COALESCE(SUM(calisma_saati), 0)
            FROM operation_batches
            WHERE operator_ismi = ? AND tarih = ? AND id <> ?
            """,
            (operator_name, work_date, int(batch_id)),
        ).fetchone()[0] or 0)
        other_work_hours = float(cur.execute(
            """
            SELECT COALESCE(SUM(calisma_saati), 0)
            FROM other_work_logs
            WHERE operator_ismi = ? AND tarih = ?
            """,
            (operator_name, work_date),
        ).fetchone()[0] or 0)
        if (
            (not pd.isna(parsed_date) and parsed_date.weekday() >= 5)
            or other_operation_hours + other_work_hours + work_hours > 9.0
        ):
            work_type = "Mesaili"

        cur.execute(
            """
            UPDATE operation_batches
            SET tarih = ?, operator_ismi = ?, calisma_tipi = ?,
                calisma_saati = ?, neden = ?, notlar = ?
            WHERE id = ?
            """,
            (
                work_date,
                operator_name,
                work_type,
                work_hours,
                str(batch_data.get("neden", "") or ""),
                str(batch_data.get("notlar", "") or ""),
                int(batch_id),
            ),
        )

        updated_count = 0
        inserted_count = 0
        submitted_existing_ids = set()
        final_keys = set()
        created_at = datetime.now().isoformat(timespec="seconds")

        for _, row in edited_rows.iterrows():
            raw_log_id = row.get("kayıt_no")
            log_id = None
            if raw_log_id is not None and not pd.isna(raw_log_id) and str(raw_log_id).strip():
                log_id = int(float(raw_log_id))
                if log_id not in allowed_ids:
                    raise ValueError("Bu kayda ait olmayan bir işlem satırı değiştirilemez.")

            oc_no = str(row.get("OC", "") or "").strip()
            pos = _pos_label(row.get("POS", ""))
            operation_name = str(row.get("İşlem", "") or "").strip()
            if not oc_no or not pos or not operation_name:
                raise ValueError("Eklenen her satırda OC, POS ve İşlem alanları doldurulmalıdır.")

            item_row = cur.execute(
                """
                SELECT id, unit_area_mm2
                FROM production_output_items
                WHERE oc_no = ? AND pos = ?
                """,
                (oc_no, pos),
            ).fetchone()
            if item_row is None:
                raise ValueError(
                    f"OC {oc_no} / {pos} üretim çıktısında bulunamadı. "
                    "Önce ilgili üretim çıktısını sisteme yüklemelisin."
                )
            item_id, unit_area_mm2 = int(item_row[0]), float(item_row[1])
            affected_item_ids.add(item_id)

            duplicate_key = (item_id, _normalize_header(operation_name))
            if duplicate_key in final_keys:
                raise ValueError(
                    f"OC {oc_no} / {pos} / {operation_name} satırı aynı kayıtta iki kez bulunamaz."
                )
            final_keys.add(duplicate_key)

            processed_qty = max(int(float(row.get("İşlem Yapılan", 0) or 0)), 0)
            fire_qty = max(int(float(row.get("Fire", 0) or 0)), 0)
            if fire_qty > processed_qty:
                raise ValueError(
                    f"OC {oc_no} / {pos}: Fire adedi işlem yapılan adetten büyük olamaz."
                )
            good_qty = processed_qty - fire_qty
            operation_hours = max(float(row.get("İşlem Süresi (Saat)", 0) or 0), 0.0)
            fire_note = str(row.get("Fire Notu", "") or "")

            if log_id is None:
                cur.execute(
                    """
                    INSERT INTO operation_work_logs (
                        batch_id, item_id, operation_name, processed_qty,
                        fire_qty, good_qty, operation_hours,
                        calculated_area_mm2, fire_note, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        int(batch_id), item_id, operation_name, processed_qty,
                        fire_qty, good_qty, operation_hours,
                        good_qty * unit_area_mm2, fire_note, created_at,
                    ),
                )
                inserted_count += 1
            else:
                submitted_existing_ids.add(log_id)
                old_item_row = cur.execute(
                    "SELECT item_id FROM operation_work_logs WHERE id = ? AND batch_id = ?",
                    (log_id, int(batch_id)),
                ).fetchone()
                if old_item_row:
                    affected_item_ids.add(int(old_item_row[0]))
                cur.execute(
                    """
                    UPDATE operation_work_logs
                    SET item_id = ?, operation_name = ?, processed_qty = ?,
                        fire_qty = ?, good_qty = ?, operation_hours = ?,
                        calculated_area_mm2 = ?, fire_note = ?
                    WHERE id = ? AND batch_id = ?
                    """,
                    (
                        item_id, operation_name, processed_qty, fire_qty, good_qty,
                        operation_hours, good_qty * unit_area_mm2, fire_note,
                        log_id, int(batch_id),
                    ),
                )
                updated_count += int(cur.rowcount or 0)

        deleted_ids = allowed_ids - submitted_existing_ids
        deleted_count = 0
        if deleted_ids:
            placeholders = ",".join("?" for _ in deleted_ids)
            cur.execute(
                f"DELETE FROM operation_work_logs WHERE batch_id = ? AND id IN ({placeholders})",
                (int(batch_id), *sorted(deleted_ids)),
            )
            deleted_count = int(cur.rowcount or 0)

        remaining_count = int(cur.execute(
            "SELECT COUNT(*) FROM operation_work_logs WHERE batch_id = ?",
            (int(batch_id),),
        ).fetchone()[0])
        if remaining_count <= 0:
            raise ValueError("Kayıtta en az bir işlem satırı kalmalıdır.")

        rebuild_operation_completion_logs(conn, list(affected_item_ids))
        conn.commit()
        _clear_data_caches()
        return {
            "updated_detail_count": updated_count,
            "inserted_detail_count": inserted_count,
            "deleted_detail_count": deleted_count,
            "affected_item_count": len(affected_item_ids),
            "work_type": work_type,
        }
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def delete_operation_batch(batch_id: int) -> dict:
    """
    Seçilen operasyon kaydını ve bağlı işlem satırlarını siler.
    POS ilerlemeleri kalan kayıtlara göre yeniden hesaplanır.
    """
    conn = get_db_connection()
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
        _clear_data_caches()
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



@st.cache_data(ttl=300, show_spinner=False)
def get_operation_batch_summary() -> pd.DataFrame:
    conn = get_db_connection()
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


@st.cache_data(ttl=300, show_spinner=False)
def get_operation_batch_details(batch_id: int) -> pd.DataFrame:
    conn = get_db_connection()
    try:
        return pd.read_sql_query(
            """
            SELECT
                w.id AS kayıt_no,
                w.item_id AS item_id,
                i.oc_no AS OC,
                i.pos AS POS,
                w.operation_name AS İşlem,
                w.operation_hours AS İşlem_Süresi_Saat,
                w.laser_plate_qty AS Lazer_Plaka_Adedi,
                w.material_type AS Malzeme,
                w.thickness_mm AS Kalınlık_mm,
                w.abkant_team_qty AS Abkant_Takım_Sayısı,
                w.abkant_teams_per_piece AS Adet_Başına_Takım,
                w.abkant_long_bend_qty AS Uzun_Kenar_Büküm_Adedi,
                w.abkant_short_bend_qty AS Kısa_Kenar_Büküm_Adedi,
                w.abkant_long_single_bend_qty AS Uzun_Kenar_Tek_Büküm,
                w.abkant_long_double_bend_qty AS Uzun_Kenar_Çift_Büküm,
                w.abkant_short_single_bend_qty AS Kısa_Kenar_Tek_Büküm,
                w.abkant_short_double_bend_qty AS Kısa_Kenar_Çift_Büküm,
                w.abkant_manual_override AS Adet_Manuel_Düzeltildi,
                w.abkant_work_mode AS Abkant_Çalışma_Şekli,
                w.abkant_coworker AS Beraber_Çalıştığı,
                w.participants_text AS Çalışma_Ekibi,
                w.piece_weight_kg AS Parça_Ağırlığı_kg,
                w.bend_type AS Büküm_Türü,
                ROUND(w.calculated_area_mm2 / 1000000.0, 1) AS Hesaplanan_Alan_m2,
                w.processed_qty AS İşlem_Yapılan,
                w.fire_qty AS Fire,
                w.good_qty AS Sağlam_İlerleyen,
                ROUND(i.unit_area_mm2 / 1000000.0, 1) AS Birim_Alan_m2,
                ROUND(w.good_qty * i.unit_area_mm2 / 1000000.0, 1) AS Sağlam_Alan_m2,
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
    worker_day_rows = expand_worker_attribution(day)

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
    m1.metric("Çalışan", int(worker_day_rows["operator_ismi"].nunique()))
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
        "m² / saat",
        f"{area_productivity / 1_000_000:.1f}",
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

    worker_summary = (
        worker_day_rows.groupby("operator_ismi", as_index=False)
        .agg(
            kayıt=("batch_id", "nunique"),
            POS=("POS", "nunique"),
            işlem=("operasyon", "nunique"),
            işlem_yapılan=("islem_yapilan", "sum"),
            sağlam_ilerleyen=("saglam_ilerleyen", "sum"),
            sağlam_alan_mm2=("saglam_alan_mm2", "sum"),
            çalışma_saati=("operasyon_saati", "sum"),
            fire=("fire", "sum"),
        )
    )
    worker_summary["adet_saat"] = (
        worker_summary["sağlam_ilerleyen"]
        / worker_summary["çalışma_saati"].replace(0, 1)
    ).round(2)
    worker_summary["Toplam Alan (m²)"] = (
        worker_summary["sağlam_alan_mm2"] / 1_000_000
    ).round(1)
    worker_summary["Günlük Alan (m²)"] = (
        worker_summary["Toplam Alan (m²)"]
    ).round(1)
    worker_summary["m2_saat"] = (
        worker_summary["Toplam Alan (m²)"]
        / worker_summary["çalışma_saati"].replace(0, 1)
    ).round(1)
    worker_summary["fire_yuzde"] = (
        worker_summary["fire"]
        / worker_summary["işlem_yapılan"].replace(0, 1)
        * 100
    ).round(2)
    worker_summary = worker_summary.drop(
        columns=["sağlam_alan_mm2"], errors="ignore"
    ).sort_values(
        "Toplam Alan (m²)",
        ascending=False,
    )

    detail = worker_day_rows[
        [
            "batch_id",
            "tarih",
            "operator_ismi",
            "ana_operator",
            "katilim_rolu",
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
    detail["Sağlam Alan (m²)"] = (
        detail["saglam_alan_mm2"] / 1_000_000
    ).round(1)
    detail = detail.drop(columns=["saglam_alan_mm2"], errors="ignore")

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
    operation_tab, other_tab, legacy_tab = st.tabs(
        ["Operasyon Kayıtları", "Diğer Çalışmalar", "Eski Sistem Kayıtları"]
    )

    with operation_tab:
        if operation_batches.empty:
            render_empty_state("Henüz operasyon kaydı yok", "İlk çalışan kaydı tamamlandığında ayrıntılar, verim ve silme seçenekleri burada görünür.", "▤")
        else:
            operation_batches = operation_batches.copy()
            operation_batches["saglam_alan_m2"] = (
                operation_batches["saglam_alan_mm2"] / 1_000_000
            ).round(1)
            operation_batches["m2_saat"] = (
                operation_batches["saglam_alan_m2"]
                / operation_batches["calisma_saati"].replace(0, 1)
            ).round(1)
            operation_batches = operation_batches.drop(
                columns=["saglam_alan_mm2"], errors="ignore"
            )
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
                    "saglam_alan_m2": "Sağlam Alan (m²)",
                    "adet_saat": "Adet/Saat",
                    "mm2_saat": "m²/Saat",
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
                "m² / saat",
                f"{float(selected_batch['mm2_saat']):.1f}",
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
            st.subheader("Operasyon Kaydını Düzenle")
            st.caption(
                "Eksik veya yanlış girilen üretim adedini, fireyi, işlem süresini "
                "ve kayıt bilgilerini silmeden değiştirebilirsin. Kaydettiğinde "
                "POS ilerlemesi otomatik olarak yeniden hesaplanır."
            )

            edit_col1, edit_col2 = st.columns(2)
            with edit_col1:
                edit_date = st.date_input(
                    "Tarih",
                    value=pd.to_datetime(selected_batch["tarih"]).date(),
                    key=f"operation_edit_date_{selected_batch_id}",
                )
                edit_worker = st.selectbox(
                    "Çalışan",
                    [name for name in WORKER_NAMES if name],
                    index=max(
                        0,
                        [name for name in WORKER_NAMES if name].index(
                            selected_batch["operator_ismi"]
                        ) if selected_batch["operator_ismi"] in [name for name in WORKER_NAMES if name] else 0,
                    ),
                    key=f"operation_edit_worker_{selected_batch_id}",
                )
                edit_type_options = ["Tam zamanlı", "Yarı zamanlı", "Mesaili"]
                edit_work_type = st.selectbox(
                    "Çalışma tipi",
                    edit_type_options,
                    index=(
                        edit_type_options.index(selected_batch["calisma_tipi"])
                        if selected_batch["calisma_tipi"] in edit_type_options else 0
                    ),
                    key=f"operation_edit_type_{selected_batch_id}",
                )
            with edit_col2:
                edit_hours = st.number_input(
                    "Toplam çalışma saati",
                    min_value=0.1,
                    max_value=24.0,
                    value=float(selected_batch["calisma_saati"]),
                    step=0.25,
                    key=f"operation_edit_hours_{selected_batch_id}",
                )
                edit_reason = st.text_input(
                    "Neden",
                    value=str(selected_batch.get("neden", "") or ""),
                    key=f"operation_edit_reason_{selected_batch_id}",
                )
                edit_note = st.text_area(
                    "Not",
                    value=str(selected_batch.get("notlar", "") or ""),
                    key=f"operation_edit_note_{selected_batch_id}",
                )

            editable_details = details.rename(
                columns={
                    "İşlem_Süresi_Saat": "İşlem Süresi (Saat)",
                    "İşlem_Yapılan": "İşlem Yapılan",
                    "Fire_Notu": "Fire Notu",
                }
            )[[
                "kayıt_no", "OC", "POS", "İşlem",
                "İşlem Yapılan", "Fire", "İşlem Süresi (Saat)", "Fire Notu"
            ]].copy()

            st.info(
                "Eksik POS eklemek için tablonun altındaki + simgesine bas. "
                "Bir satırı çıkarmak için satırı seçip çöp kutusu simgesini kullan."
            )
            all_output_items = get_production_output_summary()
            available_ocs = sorted(
                all_output_items["oc_no"].dropna().astype(str).unique().tolist()
            ) if not all_output_items.empty else []
            available_positions = sorted(
                all_output_items["pos"].dropna().astype(str).unique().tolist(),
                key=lambda value: int(re.search(r"\d+", value).group())
                if re.search(r"\d+", value) else 0,
            ) if not all_output_items.empty else []
            available_operations = sorted(
                set(details["İşlem"].dropna().astype(str).tolist())
                | set(get_performance_targets()["operation_name"].dropna().astype(str).tolist())
            )

            edited_details = st.data_editor(
                editable_details,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                disabled=["kayıt_no"],
                column_config={
                    "kayıt_no": None,
                    "OC": st.column_config.SelectboxColumn(
                        "OC", options=available_ocs, required=True
                    ),
                    "POS": st.column_config.SelectboxColumn(
                        "POS", options=available_positions, required=True
                    ),
                    "İşlem": st.column_config.SelectboxColumn(
                        "İşlem", options=available_operations, required=True
                    ),
                    "İşlem Yapılan": st.column_config.NumberColumn(
                        "İşlem Yapılan", min_value=0, step=1, format="%d", required=True
                    ),
                    "Fire": st.column_config.NumberColumn(
                        "Fire", min_value=0, step=1, format="%d", required=True
                    ),
                    "İşlem Süresi (Saat)": st.column_config.NumberColumn(
                        "İşlem Süresi (Saat)", min_value=0.0, step=0.25,
                        format="%.2f", required=True
                    ),
                },
                key=f"operation_edit_table_{selected_batch_id}",
            )

            if st.button(
                "Değişiklikleri Kaydet",
                type="primary",
                use_container_width=True,
                key=f"operation_edit_save_{selected_batch_id}",
            ):
                try:
                    result = update_operation_batch_record(
                        selected_batch_id,
                        {
                            "tarih": edit_date.isoformat(),
                            "operator_ismi": edit_worker,
                            "calisma_tipi": edit_work_type,
                            "calisma_saati": edit_hours,
                            "neden": edit_reason,
                            "notlar": edit_note,
                        },
                        edited_details,
                    )
                    st.success(
                        f"#{selected_batch_id} numaralı kayıt güncellendi. "
                        f"{result['updated_detail_count']} satır değiştirildi, "
                        f"{result['inserted_detail_count']} satır eklendi, "
                        f"{result['deleted_detail_count']} satır çıkarıldı; "
                        f"{result['affected_item_count']} POS ilerlemesi yeniden hesaplandı."
                    )
                    if result["work_type"] == "Mesaili" and edit_work_type != "Mesaili":
                        st.info(
                            "Hafta sonu veya günlük toplam 9 saati aştığı için "
                            "çalışma tipi otomatik olarak Mesaili yapıldı."
                        )
                    st.rerun()
                except Exception as exc:
                    st.error(f"Kayıt güncellenemedi: {exc}")

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

    with other_tab:
        other_logs = get_other_work_logs()
        if other_logs.empty:
            st.info("Henüz diğer çalışma kaydı yok.")
        else:
            display_other = other_logs.rename(columns={
                "id": "Kayıt No",
                "tarih": "Tarih",
                "operator_ismi": "Çalışan",
                "calisma_tipi": "Çalışma Tipi",
                "calisma_saati": "Saat",
                "is_aciklamasi": "Yapılan İş",
                "participants_text": "Çalışma Ekibi",
                "neden": "Neden",
                "notlar": "Not",
                "created_at": "Kayıt Zamanı",
            })
            st.dataframe(display_other, use_container_width=True, hide_index=True)

            other_ids = other_logs["id"].astype(int).tolist()
            selected_other_id = st.selectbox(
                "Silinecek diğer çalışma kaydı",
                other_ids,
                key="delete_other_work_select",
            )
            confirm_other_delete = st.checkbox(
                "Bu diğer çalışma kaydını silmek istediğimi onaylıyorum",
                key=f"delete_other_work_confirm_{selected_other_id}",
            )
            if st.button(
                "Seçili Diğer Çalışma Kaydını Sil",
                disabled=not confirm_other_delete,
                key=f"delete_other_work_button_{selected_other_id}",
            ):
                delete_other_work_log(selected_other_id)
                st.success("Diğer çalışma kaydı silindi.")
                st.rerun()

    with legacy_tab:
        conn = get_db_connection()
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



def _company_chart_palette() -> list[str]:
    return [
        "#123B5D", "#00A6A6", "#F59E0B", "#3B82F6",
        "#10B981", "#8B5CF6", "#EF4444", "#64748B",
    ]


def _vega_bar_chart(
    data: pd.DataFrame,
    category: str,
    value: str,
    title: str,
    value_title: str,
    horizontal: bool = True,
    color_field: str | None = None,
):
    if data.empty:
        st.info("Grafik için yeterli veri yok.")
        return

    chart_data = data.copy()
    chart_data[value] = pd.to_numeric(
        chart_data[value], errors="coerce"
    ).fillna(0)
    chart_data = chart_data.sort_values(value, ascending=False)

    color_encoding = (
        {
            "field": color_field,
            "type": "nominal",
            "legend": {"title": None, "orient": "bottom"},
            "scale": {"range": _company_chart_palette()},
        }
        if color_field and color_field in chart_data.columns
        else {"value": "#123B5D"}
    )

    tooltip = [
        {"field": category, "type": "nominal", "title": category},
        {
            "field": value,
            "type": "quantitative",
            "title": value_title,
            "format": ".1f",
        },
    ]
    if color_field and color_field in chart_data.columns and color_field != category:
        tooltip.insert(
            1,
            {"field": color_field, "type": "nominal", "title": color_field},
        )

    if horizontal:
        encoding = {
            "y": {
                "field": category,
                "type": "nominal",
                "sort": "-x",
                "title": None,
                "axis": {"labelLimit": 280, "labelPadding": 8},
            },
            "x": {
                "field": value,
                "type": "quantitative",
                "title": value_title,
                "axis": {"grid": True, "tickCount": 6},
            },
            "color": color_encoding,
            "tooltip": tooltip,
        }
        label_encoding = {
            "y": {"field": category, "type": "nominal", "sort": "-x"},
            "x": {"field": value, "type": "quantitative"},
            "text": {
                "field": value,
                "type": "quantitative",
                "format": ".1f",
            },
        }
    else:
        encoding = {
            "x": {
                "field": category,
                "type": "nominal",
                "sort": "-y",
                "title": None,
                "axis": {"labelAngle": -22, "labelLimit": 180},
            },
            "y": {
                "field": value,
                "type": "quantitative",
                "title": value_title,
                "axis": {"grid": True, "tickCount": 6},
            },
            "color": color_encoding,
            "tooltip": tooltip,
        }
        label_encoding = {
            "x": {"field": category, "type": "nominal", "sort": "-y"},
            "y": {"field": value, "type": "quantitative"},
            "text": {
                "field": value,
                "type": "quantitative",
                "format": ".1f",
            },
        }

    st.vega_lite_chart(
        chart_data,
        {
            "title": {
                "text": title,
                "anchor": "start",
                "fontSize": 18,
                "fontWeight": 700,
                "color": "#0F172A",
                "offset": 18,
            },
            "layer": [
                {
                    "mark": {
                        "type": "bar",
                        "cornerRadiusEnd": 7,
                        "size": 24,
                    },
                    "encoding": encoding,
                },
                {
                    "mark": {
                        "type": "text",
                        "align": "left" if horizontal else "center",
                        "baseline": "middle" if horizontal else "bottom",
                        "dx": 6 if horizontal else 0,
                        "dy": 0 if horizontal else -5,
                        "fontSize": 11,
                        "fontWeight": 600,
                        "color": "#334155",
                    },
                    "encoding": label_encoding,
                },
            ],
            "config": {
                "view": {"stroke": None},
                "axis": {
                    "gridColor": "#E2E8F0",
                    "domainColor": "#CBD5E1",
                    "tickColor": "#CBD5E1",
                    "labelColor": "#475569",
                    "titleColor": "#334155",
                    "labelFontSize": 12,
                    "titleFontSize": 12,
                },
                "legend": {
                    "labelColor": "#475569",
                    "symbolType": "circle",
                },
            },
            "height": max(310, min(560, 38 * len(chart_data))),
        },
        use_container_width=True,
    )


def _vega_line_chart(
    data: pd.DataFrame,
    date_column: str,
    value: str,
    title: str,
    value_title: str,
):
    if data.empty:
        st.info("Grafik için yeterli veri yok.")
        return

    chart_data = data.copy()
    chart_data[value] = pd.to_numeric(
        chart_data[value], errors="coerce"
    ).fillna(0)
    st.vega_lite_chart(
        chart_data,
        {
            "title": {
                "text": title,
                "anchor": "start",
                "fontSize": 18,
                "fontWeight": 700,
                "color": "#0F172A",
                "offset": 18,
            },
            "layer": [
                {
                    "mark": {
                        "type": "area",
                        "line": False,
                        "opacity": 0.16,
                        "color": "#00A6A6",
                    },
                    "encoding": {
                        "x": {
                            "field": date_column,
                            "type": "temporal",
                            "title": "Tarih",
                            "axis": {"format": "%d.%m", "labelAngle": 0},
                        },
                        "y": {
                            "field": value,
                            "type": "quantitative",
                            "title": value_title,
                            "axis": {"grid": True, "tickCount": 6},
                        },
                    },
                },
                {
                    "mark": {
                        "type": "line",
                        "point": {
                            "filled": True,
                            "fill": "white",
                            "stroke": "#123B5D",
                            "strokeWidth": 2,
                            "size": 70,
                        },
                        "strokeWidth": 3,
                        "color": "#123B5D",
                    },
                    "encoding": {
                        "x": {
                            "field": date_column,
                            "type": "temporal",
                            "title": "Tarih",
                            "axis": {"format": "%d.%m"},
                        },
                        "y": {
                            "field": value,
                            "type": "quantitative",
                            "title": value_title,
                        },
                        "tooltip": [
                            {
                                "field": date_column,
                                "type": "temporal",
                                "title": "Tarih",
                                "format": "%d.%m.%Y",
                            },
                            {
                                "field": value,
                                "type": "quantitative",
                                "title": value_title,
                                "format": ".1f",
                            },
                        ],
                    },
                },
            ],
            "config": {
                "view": {"stroke": None},
                "axis": {
                    "gridColor": "#E2E8F0",
                    "domainColor": "#CBD5E1",
                    "tickColor": "#CBD5E1",
                    "labelColor": "#475569",
                    "titleColor": "#334155",
                    "labelFontSize": 12,
                },
            },
            "height": 360,
        },
        use_container_width=True,
    )


def _vega_pie_chart(
    data: pd.DataFrame,
    category: str,
    value: str,
    title: str,
    value_title: str = "Alan (m²)",
    value_suffix: str = "m²",
):
    if data.empty:
        st.info("Daire grafik için yeterli veri yok.")
        return

    chart_data = data[[category, value]].copy()
    chart_data[value] = pd.to_numeric(
        chart_data[value], errors="coerce"
    ).fillna(0)
    chart_data = chart_data[chart_data[value] > 0].sort_values(
        value, ascending=False
    )
    if chart_data.empty:
        st.info("Daire grafik için yeterli veri yok.")
        return

    if len(chart_data) > 7:
        top = chart_data.head(6).copy()
        other_value = float(chart_data.iloc[6:][value].sum())
        chart_data = pd.concat(
            [
                top,
                pd.DataFrame([{category: "Diğer", value: other_value}]),
            ],
            ignore_index=True,
        )

    total_value = float(chart_data[value].sum())
    suffix_text = f" {value_suffix}" if value_suffix else ""
    st.vega_lite_chart(
        chart_data,
        {
            "title": {
                "text": title,
                "anchor": "start",
                "fontSize": 18,
                "fontWeight": 700,
                "color": "#0F172A",
                "offset": 18,
            },
            "layer": [
                {
                    "mark": {
                        "type": "arc",
                        "innerRadius": 78,
                        "outerRadius": 142,
                        "cornerRadius": 4,
                        "padAngle": 0.018,
                    },
                    "encoding": {
                        "theta": {
                            "field": value,
                            "type": "quantitative",
                            "stack": True,
                        },
                        "color": {
                            "field": category,
                            "type": "nominal",
                            "legend": {
                                "title": None,
                                "orient": "bottom",
                                "columns": 2,
                                "labelLimit": 170,
                            },
                            "scale": {"range": _company_chart_palette()},
                        },
                        "order": {
                            "field": value,
                            "type": "quantitative",
                            "sort": "descending",
                        },
                        "tooltip": [
                            {
                                "field": category,
                                "type": "nominal",
                                "title": "Kategori",
                            },
                            {
                                "field": value,
                                "type": "quantitative",
                                "title": value_title,
                                "format": ".1f",
                            },
                        ],
                    },
                },
                {
                    "mark": {
                        "type": "text",
                        "fontSize": 24,
                        "fontWeight": 700,
                        "color": "#123B5D",
                    },
                    "encoding": {
                        "text": {"value": f"{total_value:.1f}{suffix_text}"},
                    },
                },
            ],
            "config": {
                "view": {"stroke": None},
                "legend": {
                    "labelColor": "#475569",
                    "symbolType": "circle",
                },
            },
            "height": 395,
        },
        use_container_width=True,
    )


def _vega_heatmap_chart(
    data: pd.DataFrame,
    x_field: str,
    y_field: str,
    value_field: str,
    title: str,
):
    if data.empty:
        st.info("Isı haritası için yeterli veri yok.")
        return
    chart_data = data.copy()
    chart_data[value_field] = pd.to_numeric(
        chart_data[value_field], errors="coerce"
    ).fillna(0)
    st.vega_lite_chart(
        chart_data,
        {
            "title": {
                "text": title,
                "anchor": "start",
                "fontSize": 18,
                "fontWeight": 700,
                "color": "#0F172A",
                "offset": 18,
            },
            "layer": [
                {
                    "mark": {"type": "rect", "cornerRadius": 4},
                    "encoding": {
                        "x": {
                            "field": x_field,
                            "type": "nominal",
                            "title": None,
                            "axis": {"labelAngle": -25, "labelLimit": 160},
                        },
                        "y": {
                            "field": y_field,
                            "type": "nominal",
                            "title": None,
                            "axis": {"labelLimit": 230},
                        },
                        "color": {
                            "field": value_field,
                            "type": "quantitative",
                            "title": "m²",
                            "scale": {
                                "range": ["#E2E8F0", "#67E8F9", "#123B5D"]
                            },
                        },
                        "tooltip": [
                            {"field": y_field, "type": "nominal", "title": "Çalışan"},
                            {"field": x_field, "type": "nominal", "title": "Operasyon"},
                            {
                                "field": value_field,
                                "type": "quantitative",
                                "title": "Alan (m²)",
                                "format": ".1f",
                            },
                        ],
                    },
                },
                {
                    "mark": {
                        "type": "text",
                        "fontSize": 11,
                        "fontWeight": 600,
                    },
                    "encoding": {
                        "x": {"field": x_field, "type": "nominal"},
                        "y": {"field": y_field, "type": "nominal"},
                        "text": {
                            "field": value_field,
                            "type": "quantitative",
                            "format": ".1f",
                        },
                        "color": {
                            "condition": {
                                "test": f"datum['{value_field}'] > 20",
                                "value": "white",
                            },
                            "value": "#0F172A",
                        },
                    },
                },
            ],
            "config": {
                "view": {"stroke": None},
                "axis": {
                    "domain": False,
                    "ticks": False,
                    "labelColor": "#475569",
                    "labelFontSize": 12,
                },
            },
            "height": max(320, min(620, 42 * data[y_field].nunique())),
        },
        use_container_width=True,
    )


def _prepare_company_efficiency(history: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Bütün etapları kullanarak kişi-gün ve kişi bazlı m² özeti üretir."""
    if history.empty:
        return pd.DataFrame(), pd.DataFrame()

    data = expand_worker_attribution(history)
    data["tarih_dt"] = pd.to_datetime(data["tarih"], errors="coerce")
    data = data.dropna(subset=["tarih_dt"])
    data["operasyon_turu"] = data["operasyon"].map(_operation_kind)
    data["alan_m2"] = (
        pd.to_numeric(data["saglam_alan_mm2"], errors="coerce").fillna(0)
        / 1_000_000
    )
    data["boya_m2"] = data["alan_m2"].where(data["operasyon_turu"].eq("boya"), 0.0)
    data["laser_m2"] = data["alan_m2"].where(data["operasyon_turu"].eq("laser"), 0.0)
    data["saatli_alan_m2"] = data["alan_m2"].where(
        pd.to_numeric(data["operasyon_saati"], errors="coerce").fillna(0).gt(0),
        0.0,
    )
    data["cut_out_adet"] = pd.to_numeric(
        data["saglam_ilerleyen"], errors="coerce"
    ).fillna(0).where(data["operasyon_turu"].eq("laser_cut_out"), 0)

    person_day = (
        data.groupby(["tarih_dt", "operator_ismi"], as_index=False)
        .agg(
            toplam_m2=("alan_m2", "sum"),
            saatli_alan_m2=("saatli_alan_m2", "sum"),
            boya_m2=("boya_m2", "sum"),
            laser_m2=("laser_m2", "sum"),
            cut_out_adet=("cut_out_adet", "sum"),
            sağlam_adet=("saglam_ilerleyen", "sum"),
            fire=("fire", "sum"),
            operasyon_saati=("operasyon_saati", "sum"),
        )
        .sort_values(["tarih_dt", "toplam_m2"], ascending=[False, False])
    )
    for column in ["toplam_m2", "saatli_alan_m2", "boya_m2", "laser_m2", "operasyon_saati"]:
        person_day[column] = pd.to_numeric(person_day[column], errors="coerce").fillna(0).round(1)
    person_day["m2_saat"] = (
        person_day["saatli_alan_m2"]
        / person_day["operasyon_saati"].replace(0, pd.NA)
    ).fillna(0).round(1)
    person_day["cut_out_adet"] = person_day["cut_out_adet"].astype(int)

    person_summary = (
        person_day.groupby("operator_ismi", as_index=False)
        .agg(
            çalışılan_gün=("tarih_dt", "nunique"),
            toplam_m2=("toplam_m2", "sum"),
            saatli_alan_m2=("saatli_alan_m2", "sum"),
            ortalama_m2_gün=("toplam_m2", "mean"),
            toplam_boya_m2=("boya_m2", "sum"),
            ortalama_boya_m2_gün=("boya_m2", "mean"),
            cut_out_adet=("cut_out_adet", "sum"),
            fire=("fire", "sum"),
            operasyon_saati=("operasyon_saati", "sum"),
        )
        .sort_values("ortalama_m2_gün", ascending=False)
    )
    for column in [
        "toplam_m2", "ortalama_m2_gün", "toplam_boya_m2",
        "ortalama_boya_m2_gün", "operasyon_saati"
    ]:
        person_summary[column] = pd.to_numeric(
            person_summary[column], errors="coerce"
        ).fillna(0).round(1)
    person_summary["m2_saat"] = (
        person_summary["saatli_alan_m2"]
        / person_summary["operasyon_saati"].replace(0, pd.NA)
    ).fillna(0).round(1)
    person_summary["cut_out_adet"] = person_summary["cut_out_adet"].astype(int)
    return person_day, person_summary


def _daily_target_status(selected_date, actual_m2: float, target_m2: float) -> tuple[str, float]:
    if target_m2 <= 0:
        return "Hedef tanımlı değil", 0.0
    completion = actual_m2 / target_m2 * 100
    today = date.today()
    if selected_date < today:
        return ("Hedef tamamlandı" if completion >= 100 else "Hedef altında kaldı"), completion
    if selected_date > today:
        return "Henüz başlamadı", completion

    now = datetime.now()
    shift_start = 8.0
    shift_end = 18.0
    current_hour = now.hour + now.minute / 60
    elapsed_ratio = min(max((current_hour - shift_start) / (shift_end - shift_start), 0), 1)
    expected_m2 = target_m2 * elapsed_ratio
    if completion >= 100:
        return "Hedef tamamlandı", completion
    if expected_m2 <= 0:
        return "Gün başlangıcı", completion
    if actual_m2 >= expected_m2 * 1.05:
        return "Hedefin önünde", completion
    if actual_m2 < expected_m2 * 0.80:
        return "Hedefin gerisinde", completion
    return "Hedefle uyumlu", completion


def render_period_operation_pies(history: pd.DataFrame, other_logs: pd.DataFrame):
    st.markdown("### Etap bazlı işçilik süresi dağılımı")
    st.caption("Günlük, aylık veya yıllık dönemde toplam işçilik saatinin operasyonlara dağılımını gösterir.")

    if history is None:
        history = pd.DataFrame()
    data = history.copy()
    if not data.empty:
        data["tarih_dt"] = pd.to_datetime(data["tarih"], errors="coerce")
        data = data.dropna(subset=["tarih_dt"])

    period_col, date_col, project_col = st.columns(3)
    with period_col:
        period = st.radio(
            "Dönem", ["Günlük", "Aylık", "Yıllık"],
            horizontal=True, key="manager_stage_period"
        )
    with date_col:
        reference_date = st.date_input(
            "Referans tarih", value=date.today(), key="manager_stage_reference_date"
        )
    with project_col:
        project_options = []
        if not data.empty:
            project_options = sorted(
                data["proje_adi"].dropna().astype(str).unique().tolist()
            )
        if not project_options:
            st.info("Grafikleri göstermek için proje kaydı bulunamadı.")
            return
        selected_project = st.selectbox(
            "Proje", project_options, key="manager_stage_project_filter"
        )

    if not data.empty:
        if period == "Günlük":
            mask = data["tarih_dt"].dt.date == reference_date
        elif period == "Aylık":
            mask = (
                data["tarih_dt"].dt.year.eq(reference_date.year)
                & data["tarih_dt"].dt.month.eq(reference_date.month)
            )
        else:
            mask = data["tarih_dt"].dt.year.eq(reference_date.year)
        mask &= data["proje_adi"].astype(str).eq(selected_project)
        filtered = data.loc[mask].copy()
    else:
        filtered = pd.DataFrame()

    if filtered.empty:
        summary = pd.DataFrame(columns=["Etap", "Toplam Saat", "Toplam m²"])
    else:
        filtered["alan_m2"] = (
            pd.to_numeric(filtered["saglam_alan_mm2"], errors="coerce").fillna(0)
            / 1_000_000
        )
        filtered["operasyon_saati"] = pd.to_numeric(
            filtered["operasyon_saati"], errors="coerce"
        ).fillna(0)
        summary = (
            filtered.groupby("operasyon", as_index=False)
            .agg(
                **{
                    "Toplam Saat": ("operasyon_saati", "sum"),
                    "Toplam m²": ("alan_m2", "sum"),
                }
            )
            .rename(columns={"operasyon": "Etap"})
        )

    other = other_logs.copy() if isinstance(other_logs, pd.DataFrame) else pd.DataFrame()
    if not other.empty:
        other["tarih_dt"] = pd.to_datetime(other["tarih"], errors="coerce")
        other = other.dropna(subset=["tarih_dt"])
        if period == "Günlük":
            other_mask = other["tarih_dt"].dt.date == reference_date
        elif period == "Aylık":
            other_mask = (
                other["tarih_dt"].dt.year.eq(reference_date.year)
                & other["tarih_dt"].dt.month.eq(reference_date.month)
            )
        else:
            other_mask = other["tarih_dt"].dt.year.eq(reference_date.year)
        other_hours = float(
            pd.to_numeric(
                other.loc[other_mask, "calisma_saati"], errors="coerce"
            ).fillna(0).sum()
        )
        # "Diğer çalışma" kayıtlarında proje bilgisi olmadığı için proje bazlı
        # grafiğe eklenmez. Böylece seçilen projenin saatleri başka işler ile karışmaz.

    # Proje bazında toplam alan, işçilik ve m² başına işçilik özeti
    if not filtered.empty:
        project_labor = (
            filtered.groupby(["proje_adi", "OC"], as_index=False)
            .agg(**{"Toplam İşçilik (saat)": ("operasyon_saati", "sum")})
        )
        output_summary = get_production_output_summary()
        if not output_summary.empty:
            project_area = (
                output_summary.groupby(["project_name", "oc_no"], as_index=False)
                .agg(
                    toplam_alan_mm2=("requested_area_mm2", "sum"),
                    uretilen_alan_mm2=("produced_area_mm2", "sum"),
                    toplam_adet=("requested_qty", "sum"),
                    uretilen_adet=("produced_qty", "sum"),
                )
                .rename(columns={"project_name": "proje_adi", "oc_no": "OC"})
            )
            project_labor["OC"] = project_labor["OC"].astype(str)
            project_area["OC"] = project_area["OC"].astype(str)
            project_labor = project_labor.merge(project_area, on=["proje_adi", "OC"], how="left")
        else:
            project_labor["toplam_alan_mm2"] = 0.0
            project_labor["uretilen_alan_mm2"] = 0.0
            project_labor["toplam_adet"] = 0
            project_labor["uretilen_adet"] = 0

        project_labor["Proje Toplam m²"] = pd.to_numeric(project_labor["toplam_alan_mm2"], errors="coerce").fillna(0) / 1_000_000
        project_labor["Üretilen m²"] = pd.to_numeric(project_labor["uretilen_alan_mm2"], errors="coerce").fillna(0) / 1_000_000
        project_labor["Saat / m²"] = (
            project_labor["Toplam İşçilik (saat)"]
            / project_labor["Üretilen m²"].replace(0, pd.NA)
        ).fillna(0)
        project_labor["m² / Saat"] = (
            project_labor["Üretilen m²"]
            / project_labor["Toplam İşçilik (saat)"].replace(0, pd.NA)
        ).fillna(0)
        project_labor["Adet İlerleme %"] = (
            pd.to_numeric(project_labor["uretilen_adet"], errors="coerce").fillna(0)
            / pd.to_numeric(project_labor["toplam_adet"], errors="coerce").replace(0, pd.NA) * 100
        ).fillna(0)
        project_labor["m² İlerleme %"] = (
            project_labor["Üretilen m²"] / project_labor["Proje Toplam m²"].replace(0, pd.NA) * 100
        ).fillna(0)
        display_project_labor = project_labor[[
            "proje_adi", "OC", "Proje Toplam m²", "Üretilen m²",
            "Toplam İşçilik (saat)", "Saat / m²", "m² / Saat",
            "Adet İlerleme %", "m² İlerleme %",
        ]].rename(columns={"proje_adi": "Proje"})
        numeric_cols = ["Proje Toplam m²", "Üretilen m²", "Toplam İşçilik (saat)", "Saat / m²", "m² / Saat", "Adet İlerleme %", "m² İlerleme %"]
        display_project_labor[numeric_cols] = display_project_labor[numeric_cols].round(2)
        st.markdown(f"#### {selected_project} · proje özeti")
        st.dataframe(display_project_labor, use_container_width=True, hide_index=True)

    if summary.empty:
        st.info("Seçilen dönemde etap kaydı bulunamadı.")
        return

    summary["Toplam Saat"] = pd.to_numeric(
        summary["Toplam Saat"], errors="coerce"
    ).fillna(0).round(1)
    summary = summary.sort_values("Toplam Saat", ascending=False)

    # Aynı ürün her operasyondan geçtiği için etap m² değerleri toplanamaz.
    # Aksi halde bir parçanın alanı lazer, abkant, kaynak, boya vb. her etapta
    # yeniden sayılır ve proje toplamından kat kat büyük, yanıltıcı bir sonuç çıkar.
    hour_summary = summary.loc[
        summary["Toplam Saat"] > 0, ["Etap", "Toplam Saat"]
    ].copy()
    _vega_pie_chart(
        hour_summary,
        "Etap", "Toplam Saat", f"{selected_project} · {period.lower()} işçilik saati dağılımı",
        value_title="Süre (saat)", value_suffix="saat"
    )
    st.dataframe(hour_summary, use_container_width=True, hide_index=True)


def render_abkant_team_alerts(history: pd.DataFrame):
    if history is None or history.empty:
        return
    data = history.copy()
    data = data[data["operasyon"].map(_operation_kind).eq("abkant")].copy()
    if data.empty:
        return

    data["uzun_kenar_mm"] = data[["boy_mm", "en_mm"]].max(axis=1)
    data["parca_agirligi_kg"] = pd.to_numeric(
        data.get("parca_agirligi_kg", 0), errors="coerce"
    ).fillna(0)
    data["ekip_sayisi"] = data.apply(
        lambda row: 1 + len(set(
            _split_participants(row.get("beraber_calistigi", ""))
            + _split_participants(row.get("beraber_calisanlar", ""))
        )),
        axis=1,
    )
    risky = data[
        data["uzun_kenar_mm"].gt(1500)
        & data["parca_agirligi_kg"].gt(15)
        & data["ekip_sayisi"].lt(2)
    ].copy()

    st.markdown("### Abkant iki kişi çalışma kontrolü")
    if risky.empty:
        st.success("1.500 mm'den uzun ve 15 kg'dan ağır olup tek kişi kaydedilmiş abkant işi yok.")
        return

    st.error(f"{len(risky)} abkant kaydında iki kişi çalışma uyarısı var.")
    display = risky[[
        "tarih", "proje_adi", "OC", "POS", "operator_ismi",
        "uzun_kenar_mm", "parca_agirligi_kg", "abkant_calisma_sekli",
        "beraber_calistigi"
    ]].rename(columns={
        "tarih": "Tarih",
        "proje_adi": "Proje",
        "operator_ismi": "Çalışan",
        "uzun_kenar_mm": "Uzun Kenar (mm)",
        "parca_agirligi_kg": "Parça Ağırlığı (kg)",
        "abkant_calisma_sekli": "Çalışma Şekli",
        "beraber_calistigi": "Beraber Çalıştığı",
    })
    st.dataframe(display, use_container_width=True, hide_index=True)


def render_manager_efficiency_snapshot(history: pd.DataFrame):
    """Yönetici panelinde günlük hedef, işçi, m² ve çalışma saati takibi."""
    raw = history.copy() if isinstance(history, pd.DataFrame) else pd.DataFrame()
    other_logs = get_other_work_logs()

    available_dates = {date.today()}
    if not raw.empty:
        raw["tarih_dt"] = pd.to_datetime(raw["tarih"], errors="coerce")
        available_dates.update(raw["tarih_dt"].dropna().dt.date.tolist())
    if not other_logs.empty:
        other_logs = other_logs.copy()
        other_logs["tarih_dt"] = pd.to_datetime(other_logs["tarih"], errors="coerce")
        available_dates.update(other_logs["tarih_dt"].dropna().dt.date.tolist())

    st.markdown("### Günlük hedef ve çalışan takibi")
    d1, d2, d3 = st.columns([1.1, 1.0, 0.7])
    with d1:
        selected_date = st.date_input(
            "Takip tarihi",
            value=date.today(),
            min_value=min(available_dates),
            max_value=max(available_dates),
            key="manager_daily_tracking_date",
        )
    current_target = get_daily_m2_target()
    with d2:
        target_input = st.number_input(
            "Otomatik günlük m² hedefi",
            min_value=0.0,
            max_value=100000.0,
            value=float(current_target),
            step=10.0,
            key="manager_daily_m2_target_input",
            help="Bu hedef her gün otomatik olarak kullanılır. Toplam etap m²'siyle karşılaştırılır.",
        )
    with d3:
        st.write("")
        if st.button(
            "Hedefi Kaydet", use_container_width=True, key="save_daily_m2_target"
        ):
            save_daily_m2_target(target_input)
            st.success("Günlük m² hedefi kaydedildi.")
            st.rerun()

    if raw.empty:
        day_raw = pd.DataFrame()
        company_total_m2 = 0.0
        company_paint_m2 = 0.0
        person_day = pd.DataFrame()
    else:
        day_raw = raw[raw["tarih_dt"].dt.date.eq(selected_date)].copy()
        day_raw["alan_m2"] = (
            pd.to_numeric(day_raw.get("saglam_alan_mm2", 0), errors="coerce").fillna(0)
            / 1_000_000
        )
        day_raw["operasyon_turu"] = day_raw["operasyon"].map(_operation_kind)
        company_total_m2 = float(day_raw["alan_m2"].sum())
        company_paint_m2 = float(
            day_raw.loc[day_raw["operasyon_turu"].eq("boya"), "alan_m2"].sum()
        )

        person_rows = expand_worker_attribution(day_raw)
        if person_rows.empty:
            person_day = pd.DataFrame()
        else:
            person_rows["operasyon_saati"] = pd.to_numeric(
                person_rows["operasyon_saati"], errors="coerce"
            ).fillna(0)
            person_day = (
                person_rows.groupby("operator_ismi", as_index=False)
                .agg(
                    toplam_m2=("alan_m2", "sum"),
                    operasyon_saati=("operasyon_saati", "sum"),
                    sağlam_adet=("saglam_ilerleyen", "sum"),
                    fire=("fire", "sum"),
                )
            )

    day_other = pd.DataFrame()
    if not other_logs.empty:
        day_other = other_logs[other_logs["tarih_dt"].dt.date.eq(selected_date)].copy()
    if not day_other.empty:
        attributed_other = expand_other_work_attribution(day_other)
        other_worker = (
            attributed_other.groupby("operator_ismi", as_index=False)
            .agg(
                diğer_çalışma_saati=("calisma_saati", "sum"),
                diğer_iş_sayısı=("id", "count"),
            )
        )
    else:
        other_worker = pd.DataFrame(
            columns=["operator_ismi", "diğer_çalışma_saati", "diğer_iş_sayısı"]
        )

    if person_day.empty:
        person_day = other_worker.copy()
        if not person_day.empty:
            person_day["toplam_m2"] = 0.0
            person_day["operasyon_saati"] = 0.0
            person_day["sağlam_adet"] = 0
            person_day["fire"] = 0
    else:
        person_day = person_day.merge(other_worker, on="operator_ismi", how="outer")

    if not person_day.empty:
        for column in [
            "toplam_m2", "operasyon_saati", "sağlam_adet", "fire",
            "diğer_çalışma_saati", "diğer_iş_sayısı"
        ]:
            if column not in person_day.columns:
                person_day[column] = 0
            person_day[column] = pd.to_numeric(person_day[column], errors="coerce").fillna(0)
        person_day["toplam_saat"] = (
            person_day["operasyon_saati"] + person_day["diğer_çalışma_saati"]
        ).round(2)
        person_day["toplam_m2"] = person_day["toplam_m2"].round(1)
        person_day["m2_saat"] = (
            person_day["toplam_m2"]
            / person_day["operasyon_saati"].replace(0, pd.NA)
        ).fillna(0).round(1)

    target = float(get_daily_m2_target())
    status, completion_pct = _daily_target_status(
        selected_date, company_total_m2, target
    )
    remaining_m2 = max(target - company_total_m2, 0.0)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Günlük hedef", f"{target:.1f} m²")
    k2.metric("Gerçekleşen", f"{company_total_m2:.1f} m²", delta=status)
    k3.metric("Kalan", f"{remaining_m2:.1f} m²")
    k4.metric("Günlük boya", f"{company_paint_m2:.1f} m²")
    st.progress(min(max(completion_pct / 100, 0.0), 1.0))
    st.caption(
        f"{selected_date.strftime('%d.%m.%Y')} · Hedef gerçekleşme %{completion_pct:.1f}. "
        "Şirket m² toplamında her operasyon satırı bir kez sayılır; ekip çalışanları yalnızca kişi katkısında çoğaltılır."
    )

    if status in {"Hedef tamamlandı", "Hedefin önünde"}:
        st.success(f"Gidişat: {status}")
    elif status in {"Hedefin gerisinde", "Hedef altında kaldı"}:
        st.warning(f"Gidişat: {status}")
    else:
        st.info(f"Gidişat: {status}")

    if person_day.empty:
        st.info("Seçilen tarihte çalışan kaydı bulunamadı.")
        return

    left, right = st.columns([1.15, 0.85])
    with left:
        _vega_bar_chart(
            person_day.sort_values("toplam_m2", ascending=False),
            "operator_ismi", "toplam_m2",
            "Günlük çalışan bazlı m² katkısı", "Katkı alanı (m²)"
        )
    with right:
        _vega_pie_chart(
            person_day[person_day["toplam_m2"] > 0][["operator_ismi", "toplam_m2"]],
            "operator_ismi", "toplam_m2", "Günlük çalışan m² payı"
        )

    display = person_day.rename(columns={
        "operator_ismi": "Çalışan",
        "toplam_m2": "Toplam m²",
        "operasyon_saati": "Operasyon Saati",
        "diğer_çalışma_saati": "Diğer Çalışma Saati",
        "toplam_saat": "Toplam Saat",
        "m2_saat": "m²/Saat",
        "sağlam_adet": "Sağlam Adet",
        "fire": "Fire",
        "diğer_iş_sayısı": "Diğer İş Kaydı",
    })
    st.dataframe(
        display[[
            "Çalışan", "Toplam m²", "Operasyon Saati",
            "Diğer Çalışma Saati", "Toplam Saat", "m²/Saat",
            "Sağlam Adet", "Fire", "Diğer İş Kaydı"
        ]].sort_values("Toplam m²", ascending=False),
        use_container_width=True, hide_index=True,
    )

    if not day_other.empty:
        with st.expander("Günün diğer çalışma kayıtları", expanded=False):
            st.dataframe(
                day_other[[
                    "operator_ismi", "participants_text", "calisma_saati",
                    "is_aciklamasi", "calisma_tipi", "notlar"
                ]].rename(columns={
                    "operator_ismi": "Çalışan",
                    "participants_text": "Çalışma Ekibi",
                    "calisma_saati": "Saat",
                    "is_aciklamasi": "Yapılan İş",
                    "calisma_tipi": "Çalışma Tipi",
                    "notlar": "Not",
                }),
                use_container_width=True, hide_index=True,
            )


def grafikler_page():
    st.markdown(
        """
        <div style="padding:20px 22px;border-radius:16px;background:linear-gradient(135deg,#0F2F49,#145E75);color:white;margin-bottom:14px;">
            <div style="font-size:13px;letter-spacing:.08em;opacity:.78;font-weight:700;">DURLUM ÜRETİM ANALİTİĞİ</div>
            <div style="font-size:27px;font-weight:800;margin-top:4px;">Üretim ve Verimlilik Paneli</div>
            <div style="font-size:14px;opacity:.88;margin-top:5px;">Tüm etaplarda sağlam adet × POS alanı ile m² hesabı; çalışan, operasyon, boya ve lazer performansı.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    history = get_operation_history()
    if history.empty:
        render_empty_state(
            "Grafik oluşturacak operasyon kaydı yok",
            "İlk operasyon kaydı oluşturulduğunda çalışan ve etap bazlı m² analizleri burada oluşur.",
            "▥",
        )
        return

    history = history.copy()
    history["tarih_dt"] = pd.to_datetime(history["tarih"], errors="coerce")
    history = history.dropna(subset=["tarih_dt"])
    person_history = expand_worker_attribution(history)
    if history.empty:
        st.warning("Kayıtların tarih bilgileri okunamadı.")
        return

    min_date = history["tarih_dt"].min().date()
    max_date = history["tarih_dt"].max().date()
    project_rows = (
        history[["OC", "proje_adi"]]
        .dropna(subset=["OC"])
        .drop_duplicates()
        .sort_values(["proje_adi", "OC"])
    )
    project_labels = project_rows.apply(
        lambda row: f"{row['proje_adi']} — OC {row['OC']}", axis=1
    ).tolist()
    project_to_oc = dict(zip(project_labels, project_rows["OC"].astype(str)))

    with st.container(border=True):
        f0, f1, f2, f3 = st.columns([1.25, 1.35, 0.9, 0.9])
        with f0:
            selected_project_label = st.selectbox(
                "Proje / OC",
                project_labels,
                key="analysis_project_filter_v216",
            )
            selected_oc = project_to_oc[selected_project_label]
        with f1:
            selected_range = st.date_input(
                "Tarih aralığı",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                key="analysis_date_range_v216",
            )
        with f2:
            workers = ["Tümü"] + sorted(
                person_history.loc[
                    person_history["OC"].astype(str).eq(selected_oc),
                    "operator_ismi",
                ].dropna().astype(str).unique().tolist()
            )
            selected_worker = st.selectbox(
                "Çalışan",
                workers,
                key="analysis_worker_filter_v216",
            )
        with f3:
            operations = ["Tümü"] + sorted(
                history.loc[
                    history["OC"].astype(str).eq(selected_oc),
                    "operasyon",
                ].dropna().astype(str).unique().tolist()
            )
            selected_operation = st.selectbox(
                "Operasyon",
                operations,
                key="analysis_operation_filter_v216",
            )

    if isinstance(selected_range, (tuple, list)) and len(selected_range) == 2:
        start_date, end_date = selected_range
    else:
        start_date = end_date = selected_range

    base_mask = (
        history["OC"].astype(str).eq(selected_oc)
        & (history["tarih_dt"].dt.date >= start_date)
        & (history["tarih_dt"].dt.date <= end_date)
    )
    person_mask = (
        person_history["OC"].astype(str).eq(selected_oc)
        & (person_history["tarih_dt"].dt.date >= start_date)
        & (person_history["tarih_dt"].dt.date <= end_date)
    )
    if selected_operation != "Tümü":
        base_mask &= history["operasyon"].astype(str) == selected_operation
        person_mask &= person_history["operasyon"].astype(str) == selected_operation
    if selected_worker != "Tümü":
        person_mask &= person_history["operator_ismi"].astype(str) == selected_worker
        filtered = person_history.loc[person_mask].copy()
    else:
        filtered = history.loc[base_mask].copy()
    person_filtered = person_history.loc[person_mask].copy()

    if filtered.empty or person_filtered.empty:
        st.warning("Seçilen filtrelere uygun operasyon kaydı bulunamadı.")
        return

    filtered["operasyon_turu"] = filtered["operasyon"].map(_operation_kind)
    filtered["alan_m2"] = (
        pd.to_numeric(filtered["saglam_alan_mm2"], errors="coerce").fillna(0)
        / 1_000_000
    )
    filtered["alan_m2"] = filtered["alan_m2"].round(1)
    filtered["operasyon_saati"] = pd.to_numeric(
        filtered["operasyon_saati"], errors="coerce"
    ).fillna(0)
    filtered["saatli_alan_m2"] = filtered["alan_m2"].where(
        filtered["operasyon_saati"].gt(0),
        0.0,
    )
    person_filtered["operasyon_turu"] = person_filtered["operasyon"].map(_operation_kind)
    person_filtered["alan_m2"] = (
        pd.to_numeric(person_filtered["saglam_alan_mm2"], errors="coerce").fillna(0)
        / 1_000_000
    ).round(1)
    person_filtered["operasyon_saati"] = pd.to_numeric(
        person_filtered["operasyon_saati"], errors="coerce"
    ).fillna(0)
    person_filtered["saatli_alan_m2"] = person_filtered["alan_m2"].where(
        person_filtered["operasyon_saati"].gt(0), 0.0
    )

    daily = (
        filtered.groupby("tarih_dt", as_index=False)
        .agg(
            günlük_m2=("alan_m2", "sum"),
            sağlam_adet=("saglam_ilerleyen", "sum"),
            işlem_yapılan=("islem_yapilan", "sum"),
            fire=("fire", "sum"),
            çalışan=("operator_ismi", "nunique"),
        )
        .sort_values("tarih_dt")
    )
    daily["günlük_m2"] = daily["günlük_m2"].round(1)
    daily["fire_yüzde"] = (
        daily["fire"] / daily["işlem_yapılan"].replace(0, pd.NA) * 100
    ).fillna(0).round(1)

    worker_operation = (
        person_filtered.groupby(["operator_ismi", "operasyon"], as_index=False)
        .agg(alan_m2=("alan_m2", "sum"))
    )
    worker_operation["alan_m2"] = worker_operation["alan_m2"].round(1)

    worker_daily = (
        person_filtered.groupby(["tarih_dt", "operator_ismi"], as_index=False)
        .agg(
            m2=("alan_m2", "sum"),
            saatli_m2=("saatli_alan_m2", "sum"),
            sağlam_adet=("saglam_ilerleyen", "sum"),
            fire=("fire", "sum"),
            işlem_saati=("operasyon_saati", "sum"),
        )
    )
    worker_daily["m2"] = worker_daily["m2"].round(1)
    worker_summary = (
        worker_daily.groupby("operator_ismi", as_index=False)
        .agg(
            çalışılan_gün=("tarih_dt", "nunique"),
            toplam_m2=("m2", "sum"),
            saatli_toplam_m2=("saatli_m2", "sum"),
            ortalama_m2_gün=("m2", "mean"),
            toplam_saat=("işlem_saati", "sum"),
            sağlam_adet=("sağlam_adet", "sum"),
            fire=("fire", "sum"),
        )
    )
    worker_summary["toplam_m2"] = worker_summary["toplam_m2"].round(1)
    worker_summary["ortalama_m2_gün"] = worker_summary["ortalama_m2_gün"].round(1)
    worker_summary["m2_saat"] = (
        worker_summary["saatli_toplam_m2"]
        / worker_summary["toplam_saat"].replace(0, pd.NA)
    ).fillna(0).round(1)
    worker_summary = worker_summary.sort_values(
        "ortalama_m2_gün", ascending=False
    )

    operation_summary = (
        filtered.groupby("operasyon", as_index=False)
        .agg(
            alan_m2=("alan_m2", "sum"),
            saatli_alan_m2=("saatli_alan_m2", "sum"),
            sağlam_adet=("saglam_ilerleyen", "sum"),
            işlem_yapılan=("islem_yapilan", "sum"),
            fire=("fire", "sum"),
            işlem_saati=("operasyon_saati", "sum"),
            çalışan=("operator_ismi", "nunique"),
        )
    )
    operation_summary["alan_m2"] = operation_summary["alan_m2"].round(1)
    operation_summary["m2_saat"] = (
        operation_summary["saatli_alan_m2"]
        / operation_summary["işlem_saati"].replace(0, pd.NA)
    ).fillna(0).round(1)
    operation_summary["fire_yüzde"] = (
        operation_summary["fire"]
        / operation_summary["işlem_yapılan"].replace(0, pd.NA)
        * 100
    ).fillna(0).round(1)
    operation_summary = operation_summary.sort_values("alan_m2", ascending=False)

    paint = filtered[filtered["operasyon_turu"].eq("boya")].copy()
    paint_summary = pd.DataFrame()
    if not paint.empty:
        paint_daily = (
            paint.groupby(["tarih_dt", "operator_ismi"], as_index=False)
            .agg(
                boya_m2=("alan_m2", "sum"),
                boya_saati=("operasyon_saati", "sum"),
            )
        )
        paint_summary = (
            paint_daily.groupby("operator_ismi", as_index=False)
            .agg(
                çalışılan_gün=("tarih_dt", "nunique"),
                toplam_boya_m2=("boya_m2", "sum"),
                ortalama_boya_m2_gün=("boya_m2", "mean"),
                toplam_boya_saati=("boya_saati", "sum"),
            )
        )
        paint_summary["toplam_boya_m2"] = paint_summary["toplam_boya_m2"].round(1)
        paint_summary["ortalama_boya_m2_gün"] = paint_summary[
            "ortalama_boya_m2_gün"
        ].round(1)
        paint_summary["boya_m2_saat"] = (
            paint_summary["toplam_boya_m2"]
            / paint_summary["toplam_boya_saati"].replace(0, pd.NA)
        ).fillna(0).round(1)
        paint_summary = paint_summary.sort_values(
            "ortalama_boya_m2_gün", ascending=False
        )

    laser = filtered[filtered["operasyon_turu"].eq("laser")].copy()
    laser_summary = pd.DataFrame()
    if not laser.empty:
        laser["malzeme"] = laser["malzeme"].fillna("Belirtilmemiş")
        laser["kalinlik_mm"] = pd.to_numeric(
            laser["kalinlik_mm"], errors="coerce"
        ).fillna(0)
        laser_summary = (
            laser.groupby(["malzeme", "kalinlik_mm"], as_index=False)
            .agg(
                alan_m2=("alan_m2", "sum"),
                sağlam_adet=("saglam_ilerleyen", "sum"),
                plaka=("laser_plaka_adedi", "sum"),
                saat=("operasyon_saati", "sum"),
            )
        )
        laser_summary["m2_saat"] = (
            laser_summary["alan_m2"]
            / laser_summary["saat"].replace(0, pd.NA)
        ).fillna(0).round(1)
        laser_summary["adet_plaka"] = (
            laser_summary["sağlam_adet"]
            / laser_summary["plaka"].replace(0, pd.NA)
        ).fillna(0).round(1)
        laser_summary["m2_plaka"] = (
            laser_summary["alan_m2"]
            / laser_summary["plaka"].replace(0, pd.NA)
        ).fillna(0).round(1)
        laser_summary["malzeme_kalinlik"] = laser_summary.apply(
            lambda row: f"{row['malzeme']} · {row['kalinlik_mm']:.1f} mm",
            axis=1,
        )
        laser_summary["alan_m2"] = laser_summary["alan_m2"].round(1)
        laser_summary = laser_summary.sort_values("m2_saat", ascending=False)

    project_output = get_production_output_summary(selected_oc)
    if project_output.empty:
        project_total_m2 = 0.0
        project_produced_m2 = 0.0
        project_total_qty = 0
        project_produced_qty = 0
    else:
        project_total_m2 = float(project_output["requested_area_mm2"].sum()) / 1_000_000
        project_produced_m2 = float(project_output["produced_area_mm2"].sum()) / 1_000_000
        project_total_qty = int(project_output["requested_qty"].sum())
        project_produced_qty = int(project_output["produced_qty"].sum())

    total_m2 = project_produced_m2
    avg_daily_m2 = float(daily["günlük_m2"].mean()) if not daily.empty else 0.0
    worker_day_avg = float(worker_daily["m2"].mean()) if not worker_daily.empty else 0.0
    total_paint_m2 = float(paint["alan_m2"].sum()) if not paint.empty else 0.0
    total_fire = int(filtered["fire"].sum())
    total_processed = int(filtered["islem_yapilan"].sum())
    fire_rate = total_fire / total_processed * 100 if total_processed else 0.0

    progress_m2 = (project_produced_m2 / project_total_m2 * 100) if project_total_m2 else 0.0
    progress_qty = (project_produced_qty / project_total_qty * 100) if project_total_qty else 0.0

    st.caption(f"Gösterilen proje: **{selected_project_label}**")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric(
        "Proje üretilen alan",
        f"{project_produced_m2:.1f} / {project_total_m2:.1f} m²",
        delta=f"%{progress_m2:.1f}",
        delta_color="off",
    )
    k2.metric(
        "Proje üretilen adet",
        f"{project_produced_qty} / {project_total_qty}",
        delta=f"%{progress_qty:.1f}",
        delta_color="off",
    )
    k3.metric("Operasyon kayıt alanı", f"{float(filtered['alan_m2'].sum()):.1f} m²")
    k4.metric("Toplam boya", f"{total_paint_m2:.1f} m²")
    k5.metric("Fire oranı", f"%{fire_rate:.1f}", delta=f"{total_fire} adet", delta_color="inverse")

    tab_general, tab_people, tab_operations, tab_paint, tab_laser = st.tabs([
        "Genel Bakış",
        "Çalışan Verimi",
        "Operasyonlar",
        "Boya",
        "Lazer / Plaka",
    ])

    with tab_general:
        left, right = st.columns([1.45, 0.95])
        with left:
            _vega_line_chart(
                daily,
                "tarih_dt",
                "günlük_m2",
                "Günlük toplam sağlam üretim alanı",
                "Alan (m²)",
            )
        with right:
            _vega_pie_chart(
                operation_summary,
                "operasyon",
                "alan_m2",
                "Operasyonların toplam m² payı",
            )

        st.markdown("#### Günlük yönetim özeti")
        daily_display = daily.rename(columns={
            "tarih_dt": "Tarih",
            "günlük_m2": "Toplam Alan (m²)",
            "sağlam_adet": "Sağlam Adet",
            "işlem_yapılan": "İşlem Yapılan",
            "fire": "Fire",
            "çalışan": "Çalışan Sayısı",
            "fire_yüzde": "Fire (%)",
        })
        st.dataframe(
            daily_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Toplam Alan (m²)": st.column_config.NumberColumn(format="%.1f"),
                "Fire (%)": st.column_config.NumberColumn(format="%.1f"),
            },
        )

    with tab_people:
        left, right = st.columns([1.15, 1.0])
        with left:
            _vega_bar_chart(
                worker_summary,
                "operator_ismi",
                "ortalama_m2_gün",
                "Çalışan başına ortalama günlük üretim",
                "Ortalama m²/gün",
            )
        with right:
            _vega_bar_chart(
                worker_summary,
                "operator_ismi",
                "m2_saat",
                "Çalışan başına saatlik alan verimi",
                "m²/saat",
            )

        _vega_heatmap_chart(
            worker_operation,
            "operasyon",
            "operator_ismi",
            "alan_m2",
            "Çalışan × operasyon üretim alanı matrisi",
        )

        worker_display = worker_summary.drop(
            columns=["saatli_toplam_m2"], errors="ignore"
        ).rename(columns={
            "operator_ismi": "Çalışan",
            "çalışılan_gün": "Çalışılan Gün",
            "toplam_m2": "Toplam Alan (m²)",
            "ortalama_m2_gün": "Ortalama m²/Gün",
            "toplam_saat": "Toplam İşlem Saati",
            "m2_saat": "m²/Saat",
            "sağlam_adet": "Sağlam Adet",
            "fire": "Fire",
        })
        st.dataframe(
            worker_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Toplam Alan (m²)": st.column_config.NumberColumn(format="%.1f"),
                "Ortalama m²/Gün": st.column_config.NumberColumn(format="%.1f"),
                "Toplam İşlem Saati": st.column_config.NumberColumn(format="%.1f"),
                "m²/Saat": st.column_config.NumberColumn(format="%.1f"),
            },
        )

    with tab_operations:
        left, right = st.columns([1.1, 1.0])
        with left:
            _vega_bar_chart(
                operation_summary,
                "operasyon",
                "alan_m2",
                f"{selected_project_label} · Operasyon bazlı sağlam alan",
                "Alan (m²)",
            )
        with right:
            _vega_bar_chart(
                operation_summary.sort_values("m2_saat", ascending=False),
                "operasyon",
                "m2_saat",
                f"{selected_project_label} · Operasyon bazlı saatlik verim",
                "m²/saat",
            )

        _vega_bar_chart(
            operation_summary.sort_values("fire_yüzde", ascending=False),
            "operasyon",
            "fire_yüzde",
            f"{selected_project_label} · Operasyon bazlı fire oranı",
            "Fire (%)",
            horizontal=False,
        )

        operation_display = operation_summary.drop(
            columns=["saatli_alan_m2"], errors="ignore"
        ).rename(columns={
            "operasyon": "Operasyon",
            "alan_m2": "Alan (m²)",
            "sağlam_adet": "Sağlam Adet",
            "işlem_yapılan": "İşlem Yapılan",
            "fire": "Fire",
            "işlem_saati": "İşlem Saati",
            "m2_saat": "m²/Saat",
            "çalışan": "Çalışan Sayısı",
            "fire_yüzde": "Fire (%)",
        })
        st.dataframe(
            operation_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Alan (m²)": st.column_config.NumberColumn(format="%.1f"),
                "İşlem Saati": st.column_config.NumberColumn(format="%.1f"),
                "m²/Saat": st.column_config.NumberColumn(format="%.1f"),
                "Fire (%)": st.column_config.NumberColumn(format="%.1f"),
            },
        )

    with tab_paint:
        if paint_summary.empty:
            st.info("Seçilen filtrelerde boya kaydı bulunmuyor.")
        else:
            p1, p2 = st.columns(2)
            with p1:
                _vega_bar_chart(
                    paint_summary,
                    "operator_ismi",
                    "ortalama_boya_m2_gün",
                    "Boya bölümünde kişi bazlı günlük verim",
                    "Boya m²/gün",
                )
            with p2:
                _vega_bar_chart(
                    paint_summary,
                    "operator_ismi",
                    "boya_m2_saat",
                    "Boya bölümünde saatlik alan verimi",
                    "Boya m²/saat",
                )
            paint_display = paint_summary.rename(columns={
                "operator_ismi": "Çalışan",
                "çalışılan_gün": "Çalışılan Gün",
                "toplam_boya_m2": "Toplam Boya Alanı (m²)",
                "ortalama_boya_m2_gün": "Ortalama Boya m²/Gün",
                "toplam_boya_saati": "Toplam Boya Saati",
                "boya_m2_saat": "Boya m²/Saat",
            })
            st.dataframe(
                paint_display,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Toplam Boya Alanı (m²)": st.column_config.NumberColumn(format="%.1f"),
                    "Ortalama Boya m²/Gün": st.column_config.NumberColumn(format="%.1f"),
                    "Toplam Boya Saati": st.column_config.NumberColumn(format="%.1f"),
                    "Boya m²/Saat": st.column_config.NumberColumn(format="%.1f"),
                },
            )

    with tab_laser:
        if laser_summary.empty:
            st.info("Seçilen filtrelerde malzeme ve kalınlık bilgili lazer kaydı bulunmuyor.")
        else:
            l1, l2 = st.columns(2)
            with l1:
                _vega_bar_chart(
                    laser_summary,
                    "malzeme_kalinlik",
                    "m2_saat",
                    "Malzeme ve kalınlığa göre lazer alan verimi",
                    "m²/saat",
                    color_field="malzeme",
                )
            with l2:
                _vega_bar_chart(
                    laser_summary,
                    "malzeme_kalinlik",
                    "adet_plaka",
                    "Malzeme ve kalınlığa göre plaka verimi",
                    "Adet/plaka",
                    color_field="malzeme",
                )
            laser_display = laser_summary.rename(columns={
                "malzeme": "Malzeme",
                "kalinlik_mm": "Kalınlık (mm)",
                "alan_m2": "Toplam Alan (m²)",
                "sağlam_adet": "Sağlam Adet",
                "plaka": "Plaka Adedi",
                "saat": "İşlem Saati",
                "m2_saat": "m²/Saat",
                "adet_plaka": "Adet/Plaka",
                "m2_plaka": "m²/Plaka",
            })
            st.dataframe(
                laser_display[[
                    "Malzeme", "Kalınlık (mm)", "Toplam Alan (m²)",
                    "Sağlam Adet", "Plaka Adedi", "İşlem Saati",
                    "m²/Saat", "Adet/Plaka", "m²/Plaka",
                ]],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Kalınlık (mm)": st.column_config.NumberColumn(format="%.1f"),
                    "Toplam Alan (m²)": st.column_config.NumberColumn(format="%.1f"),
                    "İşlem Saati": st.column_config.NumberColumn(format="%.1f"),
                    "m²/Saat": st.column_config.NumberColumn(format="%.1f"),
                    "Adet/Plaka": st.column_config.NumberColumn(format="%.1f"),
                    "m²/Plaka": st.column_config.NumberColumn(format="%.1f"),
                },
            )


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

    conn = get_db_connection()
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
    other_work_logs = get_other_work_logs()
    render_manager_efficiency_snapshot(operation_history)
    st.divider()
    render_period_operation_pies(operation_history, other_work_logs)
    st.divider()
    render_abkant_team_alerts(operation_history)
    st.divider()
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

        worker_attribution = expand_worker_attribution(operation_history)
        operation_worker_summary = (
            worker_attribution.groupby("operator_ismi", as_index=False)
            .agg(
                islem_satiri=("batch_id", "count"),
                islem_yapilan=("islem_yapilan", "sum"),
                saglam_ilerleyen=("saglam_ilerleyen", "sum"),
                saglam_alan_mm2=("saglam_alan_mm2", "sum"),
                calisilan_gun=("tarih", "nunique"),
                calisma_saati=("operasyon_saati", "sum"),
                fire=("fire", "sum"),
            )
        )
        operation_worker_summary["toplam_m2"] = (
            operation_worker_summary["saglam_alan_mm2"] / 1_000_000
        ).round(1)
        operation_worker_summary["m2_gun"] = (
            operation_worker_summary["toplam_m2"]
            / operation_worker_summary["calisilan_gun"].replace(0, 1)
        ).round(1)
        operation_worker_summary = operation_worker_summary.drop(
            columns=["saglam_alan_mm2"], errors="ignore"
        ).rename(columns={
            "operator_ismi": "İşçi",
            "islem_satiri": "İşlem Satırı",
            "islem_yapilan": "İşlem Yapılan",
            "saglam_ilerleyen": "Sağlam İlerleyen",
            "calisilan_gun": "Çalışılan Gün",
            "toplam_m2": "Toplam Alan (m²)",
            "m2_gun": "Ortalama m²/Gün",
            "fire": "Fire",
            "calisma_saati": "Çalışma Saati",
        })
        st.caption("Abkant ve boya ekip arkadaşları kişi bazlı katkıda ayrı ayrı görünür. Şirket toplamında aynı iş yalnızca bir kez sayılır.")
        st.dataframe(operation_worker_summary, use_container_width=True, hide_index=True)

        st.markdown("#### Lazer kalınlık ve plaka verimi")
        laser_history = operation_history[
            operation_history["malzeme"].fillna("").astype(str).str.strip().ne("")
        ].copy()
        if laser_history.empty:
            st.info("Malzeme bilgili lazer kaydı oluştuğunda burada m²/saat karşılaştırması görünecek.")
        else:
            material_options = ["Tümü"] + sorted(
                laser_history["malzeme"].dropna().astype(str).unique().tolist()
            )
            selected_material = st.selectbox(
                "Malzemeye göre verim",
                material_options,
                key="manager_material_efficiency_filter",
            )
            if selected_material != "Tümü":
                laser_history = laser_history[
                    laser_history["malzeme"].astype(str) == selected_material
                ]
            material_efficiency = (
                laser_history.groupby(
                    ["malzeme", "kalinlik_mm", "boy_mm", "en_mm"],
                    as_index=False,
                )
                .agg(
                    kayıt=("batch_id", "nunique"),
                    plaka_adedi=("laser_plaka_adedi", "sum"),
                    sağlam_adet=("saglam_ilerleyen", "sum"),
                    sağlam_alan_mm2=("saglam_alan_mm2", "sum"),
                    işlem_saati=("operasyon_saati", "sum"),
                )
            )
            material_efficiency["sağlam_alan_m2"] = (
                material_efficiency["sağlam_alan_mm2"] / 1_000_000
            ).round(1)
            safe_hours = material_efficiency["işlem_saati"].replace(0, pd.NA)
            safe_plates = material_efficiency["plaka_adedi"].replace(0, pd.NA)
            material_efficiency["adet_saat"] = (
                material_efficiency["sağlam_adet"] / safe_hours
            ).fillna(0).round(1)
            material_efficiency["plaka_saat"] = (
                material_efficiency["plaka_adedi"] / safe_hours
            ).fillna(0).round(1)
            material_efficiency["adet_plaka"] = (
                material_efficiency["sağlam_adet"] / safe_plates
            ).fillna(0).round(1)
            material_efficiency["m2_plaka"] = (
                material_efficiency["sağlam_alan_m2"] / safe_plates
            ).fillna(0).round(1)
            material_efficiency["m2_saat"] = (
                material_efficiency["sağlam_alan_m2"] / safe_hours
            ).fillna(0).round(1)
            material_efficiency["ölçü"] = material_efficiency.apply(
                lambda row: f"{row['boy_mm']:.0f} × {row['en_mm']:.0f} mm",
                axis=1,
            )

            thickness_efficiency = (
                laser_history.groupby(
                    ["malzeme", "kalinlik_mm"],
                    as_index=False,
                )
                .agg(
                    kayıt=("batch_id", "nunique"),
                    plaka_adedi=("laser_plaka_adedi", "sum"),
                    sağlam_adet=("saglam_ilerleyen", "sum"),
                    sağlam_alan_mm2=("saglam_alan_mm2", "sum"),
                    işlem_saati=("operasyon_saati", "sum"),
                )
            )
            thickness_efficiency["sağlam_alan_m2"] = (
                thickness_efficiency["sağlam_alan_mm2"] / 1_000_000
            ).round(1)
            thickness_hours = thickness_efficiency["işlem_saati"].replace(0, pd.NA)
            thickness_plates = thickness_efficiency["plaka_adedi"].replace(0, pd.NA)
            thickness_efficiency["adet_saat"] = (
                thickness_efficiency["sağlam_adet"] / thickness_hours
            ).fillna(0).round(1)
            thickness_efficiency["m2_saat"] = (
                thickness_efficiency["sağlam_alan_m2"] / thickness_hours
            ).fillna(0).round(1)
            thickness_efficiency["adet_plaka"] = (
                thickness_efficiency["sağlam_adet"] / thickness_plates
            ).fillna(0).round(1)
            thickness_efficiency["m2_plaka"] = (
                thickness_efficiency["sağlam_alan_m2"] / thickness_plates
            ).fillna(0).round(1)
            thickness_efficiency = thickness_efficiency.sort_values(
                "m2_saat",
                ascending=False,
            )

            st.markdown("##### Kalınlığa göre genel lazer verimi")
            st.dataframe(
                thickness_efficiency[
                    [
                        "malzeme", "kalinlik_mm", "kayıt", "plaka_adedi",
                        "sağlam_adet", "işlem_saati", "adet_saat",
                        "sağlam_alan_m2", "m2_saat", "adet_plaka", "m2_plaka",
                    ]
                ].rename(
                    columns={
                        "malzeme": "Malzeme",
                        "kalinlik_mm": "Kalınlık (mm)",
                        "kayıt": "Kayıt",
                        "plaka_adedi": "Plaka",
                        "sağlam_adet": "Sağlam Adet",
                        "işlem_saati": "İşlem Saati",
                        "adet_saat": "Adet/Saat",
                        "sağlam_alan_m2": "Alan (m²)",
                        "m2_saat": "m²/Saat",
                        "adet_plaka": "Adet/Plaka",
                        "m2_plaka": "m²/Plaka",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

            material_efficiency = material_efficiency.sort_values(
                "m2_saat",
                ascending=False,
            )
            st.markdown("##### Kalınlık ve parça ölçüsüne göre ayrıntı")
            st.dataframe(
                material_efficiency[
                    [
                        "malzeme", "kalinlik_mm", "ölçü", "kayıt",
                        "plaka_adedi", "sağlam_adet", "işlem_saati",
                        "adet_saat", "sağlam_alan_m2", "m2_saat",
                        "adet_plaka", "m2_plaka", "plaka_saat",
                    ]
                ].rename(
                    columns={
                        "malzeme": "Malzeme",
                        "kalinlik_mm": "Kalınlık (mm)",
                        "ölçü": "Boy × En",
                        "kayıt": "Kayıt",
                        "plaka_adedi": "Plaka",
                        "sağlam_adet": "Sağlam Adet",
                        "işlem_saati": "İşlem Saati",
                        "adet_saat": "Adet/Saat",
                        "sağlam_alan_m2": "Sağlam Alan (m²)",
                        "m2_saat": "m²/Saat",
                        "adet_plaka": "Adet/Plaka",
                        "m2_plaka": "m²/Plaka",
                        "plaka_saat": "Plaka/Saat",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )
            best_speed_row = material_efficiency.iloc[0]
            best_plate_row = material_efficiency.sort_values(
                "adet_plaka",
                ascending=False,
            ).iloc[0]
            st.success(
                f"En yüksek alan verimi: {best_speed_row['malzeme']} · "
                f"{best_speed_row['ölçü']} · {best_speed_row['kalinlik_mm']:.1f} mm · "
                f"{best_speed_row['m2_saat']:.1f} m²/saat"
            )
            st.info(
                f"Plaka başına en yüksek sağlam adet: {best_plate_row['malzeme']} · "
                f"{best_plate_row['ölçü']} · {best_plate_row['kalinlik_mm']:.1f} mm · "
                f"{best_plate_row['adet_plaka']:.1f} adet/plaka"
            )

        area_operations = operation_history[
            operation_history["operasyon"].map(_operation_kind).isin(
                ["kaynak", "cita", "boya"]
            )
        ].copy()
        if not area_operations.empty:
            st.markdown("#### Kaynak, çıta ve boya alan verimi")
            area_efficiency = (
                area_operations.groupby(
                    ["operasyon", "boy_mm", "en_mm"],
                    as_index=False,
                )
                .agg(
                    sağlam_alan_mm2=("saglam_alan_mm2", "sum"),
                    işlem_saati=("operasyon_saati", "sum"),
                    sağlam_adet=("saglam_ilerleyen", "sum"),
                )
            )
            area_efficiency["Alan (m²)"] = (
                area_efficiency["sağlam_alan_mm2"] / 1_000_000
            ).round(1)
            area_efficiency["Verim (m²/saat)"] = (
                area_efficiency["Alan (m²)"]
                / area_efficiency["işlem_saati"].replace(0, 1)
            ).round(1)
            area_efficiency["Boy × En"] = area_efficiency.apply(
                lambda row: f"{row['boy_mm']:.0f} × {row['en_mm']:.0f} mm",
                axis=1,
            )
            st.dataframe(
                area_efficiency[
                    [
                        "operasyon", "Boy × En", "sağlam_adet",
                        "Alan (m²)", "işlem_saati", "Verim (m²/saat)",
                    ]
                ].rename(
                    columns={
                        "operasyon": "İşlem",
                        "sağlam_adet": "Sağlam Adet",
                        "işlem_saati": "İşlem Saati",
                    }
                ).sort_values("Verim (m²/saat)", ascending=False),
                use_container_width=True,
                hide_index=True,
            )

        st.caption("Ayrıntılı POS ve etap ilerlemesi için Operasyon Takibi sayfasını kullan.")
        st.divider()

    if not other_work_logs.empty:
        st.markdown("### Diğer çalışma takibi")
        other_display = other_work_logs.copy()
        other_display["tarih_dt"] = pd.to_datetime(other_display["tarih"], errors="coerce")
        o1, o2 = st.columns(2)
        o1.metric("Diğer çalışma kaydı", int(len(other_display)))
        o2.metric("Toplam diğer çalışma saati", f"{float(other_display['calisma_saati'].sum()):.1f} saat")
        st.dataframe(
            other_display[[
                "tarih", "operator_ismi", "participants_text",
                "is_aciklamasi", "calisma_saati", "calisma_tipi", "notlar"
            ]].rename(columns={
                "tarih": "Tarih",
                "operator_ismi": "Çalışan",
                "participants_text": "Çalışma Ekibi",
                "is_aciklamasi": "Yapılan İş",
                "calisma_saati": "Saat",
                "calisma_tipi": "Çalışma Tipi",
                "notlar": "Not",
            }),
            use_container_width=True, hide_index=True,
        )
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

    conn = get_db_connection()
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
        total_remaining = int(overview["remaining_qty"].sum())
        total_overproduction = int(overview["overproduction_qty"].sum())

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Toplam istenen", total_requested)
        c2.metric("Tam biten", total_completed)
        c3.metric("Fazla üretim", total_overproduction)
        c4.metric("Üretimde olan", total_in_production)
        c5.metric("Kalan", total_remaining)

        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Operasyon firesi", total_fire)
        d2.metric("OC sayısı", int(overview["oc_no"].nunique()))
        d3.metric("POS sayısı", int(overview["item_id"].nunique()))
        completion_rate = (
            total_completed / total_requested * 100
            if total_requested
            else 0.0
        )
        d4.metric("Tamamlanma", f"%{completion_rate:.1f}")

        summary_table = overview[
            [
                "oc_no",
                "project_name",
                "pos",
                "requested_qty",
                "produced_qty",
                "overproduction_qty",
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
                "overproduction_qty": "Fazla Üretim",
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

    conn = get_db_connection()
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
