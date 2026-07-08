import sqlite3
import unicodedata
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
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
# UYGULAMA
# -----------------------------
def main():
    st.set_page_config(page_title="POS Üretim Takip", layout="wide")
    init_db()

    st.title("POS Üretim ve Fire Takip Uygulaması")
    st.caption("İşçi çalışma süresi, POS, kombinasyon, adet ve fire bilgilerini kaydeder. — Sürüm 1.8 TAM")
    st.success("GÜNCEL SÜRÜM 1.8: Üretime Devam Et, Güncelleme Geçmişi ve önceki tüm özellikler aktiftir.")

    with st.sidebar:
        st.header("Menü")
        st.info("Sürüm 1.8 TAM")
        page = st.radio("Sayfa", ["Yeni Kayıt", "Üretime Devam Et", "Kayıtlar", "Grafikler", "Yönetici Paneli", "Gün Sonu Kontrolü", "Özet"], index=0)
        st.divider()
        uploaded_excel = st.file_uploader("Ayar Excel'ini değiştir", type=["xlsx"], help="Boş bırakırsan uygulama klasöründeki ayar_dosyasi.xlsx kullanılır.")

    (reasons, combinations), source_info = load_config(uploaded_excel)
    st.sidebar.caption(source_info)

    if page == "Yeni Kayıt":
        if not combinations:
            st.info("Devam etmek için ayar Excel'i yükle veya uygulama klasörüne ayar_dosyasi.xlsx koy.")
            st.stop()
        yeni_kayit_page(reasons, combinations)
    elif page == "Üretime Devam Et":
        if not combinations:
            st.info("Devam etmek için ayar Excel'i yükle veya uygulama klasörüne ayar_dosyasi.xlsx koy.")
            st.stop()
        uretime_devam_page(combinations)
    elif page == "Kayıtlar":
        kayitlar_page()
    elif page == "Grafikler":
        grafikler_page()
    elif page == "Yönetici Paneli":
        yonetici_paneli_page()
    elif page == "Gün Sonu Kontrolü":
        gun_sonu_kontrolu_page()
    else:
        ozet_page()


def yeni_kayit_page(reasons, combinations):
    pos_list = [f"POS{i}" for i in range(10, 501, 10)]
    combo_names = [c["ad"] for c in combinations]

    st.subheader("Yeni Üretim Kaydı")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        tarih = st.date_input("Tarih", value=date.today())
    with col2:
        operator_ismi = st.text_input("Operatör / İşçi ismi")
    with col3:
        proje = st.text_input("Proje")
    with col4:
        pos = st.selectbox("POS", pos_list)

    col5, col6 = st.columns(2)
    with col5:
        kombinasyon_adi = st.selectbox("Kombinasyon", combo_names)
    selected_combo = next(c for c in combinations if c["ad"] == kombinasyon_adi)
    operations = selected_combo["operasyonlar"]

    with col6:
        st.write("Seçilen operasyonlar")
        st.write(" → ".join(operations))

    st.divider()

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        calisma_tipi = st.selectbox("Çalışma tipi", ["Tam zamanlı", "Yarı zamanlı"])
    with c2:
        if calisma_tipi == "Tam zamanlı":
            calisma_saati = st.number_input("Çalışma saati", min_value=0.0, max_value=24.0, value=9.0, step=0.5)
        else:
            calisma_saati = st.number_input("Çalışma saati", min_value=0.0, max_value=24.0, value=6.0, step=0.5)
    with c3:
        neden = st.selectbox("Neden", [""] + reasons, disabled=(calisma_tipi == "Tam zamanlı"))
    with c4:
        siparis_adedi = st.number_input("Sipariş adedi", min_value=0, value=6, step=1)
    with c5:
        fire_adedi = st.number_input("Fire adedi", min_value=0, value=0, step=1)

    c6, c7, c8 = st.columns(3)
    with c6:
        saglam_tamamlanan = st.number_input("Sağlam tamamlanan adet", min_value=0, value=min(4, siparis_adedi), step=1)
    with c7:
        uretim_yuku = int(siparis_adedi + fire_adedi)
        st.metric("Üretim yükü", f"{uretim_yuku} adet", help="Sipariş adedi + fire adedi")
    with c8:
        kalan = max(siparis_adedi - saglam_tamamlanan, 0)
        st.metric("Kalan sağlam ihtiyaç", f"{kalan} adet")

    notlar = st.text_area("Genel not")

    st.divider()
    st.subheader("Adet Bazlı Operasyon Takibi")
    st.caption("Örneğin 6 sipariş + 1 fire varsa burada 7 adet görünür. Her adet için hangi operasyonların yapıldığını işaretle.")

    operation_rows = []
    fire_options = [""] + operations
    takip_edilen_urun_sayisi = max(uretim_yuku, 1)

    st.checkbox(
        "Tüm ürünlerde tüm operasyonları seç",
        key="select_all_all_units",
        on_change=set_all_operations_for_all_units,
        args=(takip_edilen_urun_sayisi, len(operations)),
        help="Bütün ürünlerin bütün operasyonları tamamlandıysa tek tıkla hepsini işaretler. İşareti kaldırırsan hepsini kaldırır.",
    )

    for urun_no in range(1, takip_edilen_urun_sayisi + 1):
        with st.expander(f"Ürün / Adet {urun_no}", expanded=(urun_no <= 2)):
            st.checkbox(
                "Bu ürün için tüm operasyonları seç",
                key=f"select_all_{urun_no}",
                on_change=set_all_operations_for_unit,
                args=(urun_no, len(operations)),
                help="Bu üründeki bütün operasyonları tek seferde seçer veya kaldırır.",
            )

            top_cols = st.columns([1, 2, 3])
            with top_cols[0]:
                fire_var = st.checkbox("Bu adette fire var", key=f"fire_{urun_no}")
            with top_cols[1]:
                fire_operasyonu = st.selectbox("Fire hangi operasyonda oldu?", fire_options, key=f"fireop_{urun_no}", disabled=not fire_var)
            with top_cols[2]:
                fire_notu = st.text_input("Fire notu", key=f"firenote_{urun_no}", disabled=not fire_var)

            cols = st.columns(min(max(len(operations), 1), 5))
            for idx, op in enumerate(operations, start=1):
                with cols[(idx - 1) % len(cols)]:
                    done = st.checkbox(op, key=f"op_{urun_no}_{idx}")

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

    if st.button("Kaydı Veritabanına Ekle", type="primary"):
        errors = []
        if not operator_ismi.strip():
            errors.append("Operatör / işçi ismi boş olamaz.")
        if siparis_adedi <= 0:
            errors.append("Sipariş adedi 0'dan büyük olmalı.")
        if calisma_tipi == "Tam zamanlı" and calisma_saati != 9:
            st.warning("Tam zamanlı çalışma için varsayılan süre 9 saattir. Yine de farklı süreyle kayıt yapılabilir.")
        if calisma_tipi == "Yarı zamanlı" and not neden:
            errors.append("Yarı zamanlı çalışma için neden seçilmeli.")
        if saglam_tamamlanan > siparis_adedi:
            errors.append("Sağlam tamamlanan adet sipariş adedinden büyük olamaz.")

        fire_marked_units = len({row["urun_no"] for row in operation_rows if row["fire_var"]})
        if fire_marked_units != fire_adedi:
            errors.append(f"Fire adedi {fire_adedi} girildi ama adet bazlı işaretlenen fire sayısı {fire_marked_units}.")

        if errors:
            for err in errors:
                st.error(err)
            st.stop()

        session_data = {
            "tarih": str(tarih),
            "operator_ismi": operator_ismi.strip(),
            "proje": proje.strip(),
            "pos": pos,
            "kombinasyon_adi": kombinasyon_adi,
            "calisma_tipi": calisma_tipi,
            "calisma_saati": float(calisma_saati),
            "neden": neden,
            "siparis_adedi": int(siparis_adedi),
            "saglam_tamamlanan": int(saglam_tamamlanan),
            "fire_adedi": int(fire_adedi),
            "uretim_yuku": int(uretim_yuku),
            "notlar": notlar,
        }

        session_id = insert_session(session_data, operation_rows)
        st.success(f"Kayıt eklendi. Kayıt no: {session_id}")




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
        f"Proje: {selected_session['proje'] or '-'} | "
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
    st.dataframe(sessions, use_container_width=True)

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
    csv = sessions.to_csv(index=False).encode("utf-8-sig")
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
        filtered_sessions.to_excel(writer, sheet_name="Filtreli Kayıtlar", index=False)
        by_operator.to_excel(writer, sheet_name="İşçi Özeti", index=False)
        by_pos.to_excel(writer, sheet_name="POS Özeti", index=False)
        daily.to_excel(writer, sheet_name="Günlük Özet", index=False)
        fire_by_operation.to_excel(writer, sheet_name="Fire Operasyonları", index=False)
    return output.getvalue()


def yonetici_paneli_page():
    st.subheader("Yönetici Paneli")
    st.caption(
        "Tarih, işçi, proje ve POS filtreleriyle üretim, fire, çalışma süresi ve verimlilik durumunu gösterir."
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
            "Proje",
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

    work = df.head(max_rows).copy()
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
        sessions_day.to_excel(writer, sheet_name="Günlük Kayıtlar", index=False)
        worker_summary.to_excel(writer, sheet_name="İşçi Özeti", index=False)
        pos_summary.to_excel(writer, sheet_name="POS Özeti", index=False)
        incomplete_ops.to_excel(writer, sheet_name="Eksik Operasyonlar", index=False)
        fire_details.to_excel(writer, sheet_name="Fire Ayrıntıları", index=False)
        part_time.to_excel(writer, sheet_name="Yarı Zamanlı", index=False)
        untouched_units.to_excel(writer, sheet_name="İşlemsiz Ürünler", index=False)
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
        st.dataframe(pos_summary, use_container_width=True)
        st.bar_chart(pos_summary.set_index("pos")[["sipariş", "tamamlanan", "fire"]])

    with tabs[1]:
        if incomplete_ops.empty:
            st.success("Eksik operasyon bulunmuyor.")
        else:
            st.dataframe(incomplete_ops, use_container_width=True)

    with tabs[2]:
        if fire_details.empty:
            st.success("Fire kaydı bulunmuyor.")
        else:
            st.dataframe(fire_details, use_container_width=True)

    with tabs[3]:
        st.write("İşçi özeti")
        st.dataframe(worker_summary, use_container_width=True)
        st.write("Yarı zamanlı çalışanlar")
        if part_time.empty:
            st.success("Yarı zamanlı çalışma kaydı bulunmuyor.")
        else:
            st.dataframe(part_time, use_container_width=True)

    with tabs[4]:
        if untouched_units.empty:
            st.success("Operasyon işareti olmayan ürün bulunmuyor.")
        else:
            st.dataframe(untouched_units, use_container_width=True)

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
