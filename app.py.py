import sqlite3
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

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



def delete_session(session_id: int):
    """Seçilen ana kaydı ve ona bağlı operasyon detaylarını siler."""
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
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
# UYGULAMA
# -----------------------------
def main():
    st.set_page_config(page_title="POS Üretim Takip", layout="wide")
    init_db()

    st.title("POS Üretim ve Fire Takip Uygulaması")
    st.caption("İşçi çalışma süresi, POS, kombinasyon, adet ve fire bilgilerini kaydeder.")

    with st.sidebar:
        st.header("Menü")
        page = st.radio("Sayfa", ["Yeni Kayıt", "Kayıtlar", "Özet"], index=0)
        st.divider()
        uploaded_excel = st.file_uploader("Ayar Excel'ini değiştir", type=["xlsx"], help="Boş bırakırsan uygulama klasöründeki ayar_dosyasi.xlsx kullanılır.")

    (reasons, combinations), source_info = load_config(uploaded_excel)
    st.sidebar.caption(source_info)

    if page == "Yeni Kayıt":
        if not combinations:
            st.info("Devam etmek için ayar Excel'i yükle veya uygulama klasörüne ayar_dosyasi.xlsx koy.")
            st.stop()
        yeni_kayit_page(reasons, combinations)
    elif page == "Kayıtlar":
        kayitlar_page()
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

    for urun_no in range(1, max(uretim_yuku, 1) + 1):
        with st.expander(f"Ürün / Adet {urun_no}", expanded=(urun_no <= 2)):
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


def kayitlar_page():
    st.subheader("Kayıtlar")
    conn = sqlite3.connect(DB_PATH)
    sessions = pd.read_sql_query("SELECT * FROM work_sessions ORDER BY id DESC", conn)

    if sessions.empty:
        st.info("Henüz kayıt yok.")
        conn.close()
        return

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
    toplam_urun = int(selected_session["uretim_yuku"])

    st.divider()
    st.subheader("Kombinasyon İlerleme Özeti")

    combo_summary = (
        details.groupby("operasyon_adi", as_index=False)["yapildi"]
        .sum()
        .rename(columns={"yapildi": "yapilan_adet"})
    )
    combo_summary["toplam_urun"] = toplam_urun
    combo_summary["kalan_adet"] = (combo_summary["toplam_urun"] - combo_summary["yapilan_adet"]).clip(lower=0)
    combo_summary["tamamlanma_yuzdesi"] = (
        (combo_summary["yapilan_adet"] / combo_summary["toplam_urun"].replace(0, 1)) * 100
    ).round(1)

    st.caption(
        f"Seçilen kayıt: {selected_session['pos']} / {selected_session['kombinasyon_adi']} - Toplam takip edilen adet: {toplam_urun}"
    )
    st.dataframe(combo_summary, use_container_width=True)

    grafik_operasyon = st.selectbox(
        "Daire grafiğinde görmek istediğin operasyon",
        combo_summary["operasyon_adi"].tolist(),
        key=f"pie_select_{selected_id}",
    )

    selected_op = combo_summary[combo_summary["operasyon_adi"] == grafik_operasyon].iloc[0]
    yapilan = int(selected_op["yapilan_adet"])
    kalan = int(selected_op["kalan_adet"])

    gcol1, gcol2 = st.columns([1, 1])
    with gcol1:
        fig, ax = plt.subplots(figsize=(4.5, 4.5))
        values = [yapilan, kalan]
        labels = ["Yapılan", "Kalan"]

        if sum(values) == 0:
            ax.text(0.5, 0.5, "Veri yok", ha="center", va="center", fontsize=14)
            ax.axis("off")
        else:
            ax.pie(
                values,
                labels=labels,
                autopct="%1.1f%%",
                startangle=90,
                wedgeprops={"width": 0.55, "edgecolor": "white"},
            )
            ax.set_title(f"{grafik_operasyon} operasyonu")
            ax.axis("equal")

        st.pyplot(fig)
        plt.close(fig)

    with gcol2:
        st.metric("Yapılan", f"{yapilan} adet")
        st.metric("Kalan", f"{kalan} adet")
        st.metric("Tamamlanma", f"%{float(selected_op['tamamlanma_yuzdesi']):.1f}")
        st.info(
            "Bu grafik, seçilen operasyonda toplam takip edilen adet içinden kaç tanesinin yapıldığını ve kaç tanesinin kaldığını gösterir."
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
