import sqlite3
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

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


# -----------------------------
# UYGULAMA
# -----------------------------
def main():
    st.set_page_config(page_title="POS Üretim Takip", layout="wide")
    init_db()

    st.title("POS Üretim ve Fire Takip Uygulaması")
    st.caption("İşçi çalışma süresi, POS, kombinasyon, adet ve fire bilgilerini kaydeder. — Sürüm 1.6 TAM")
    st.success("GÜNCEL SÜRÜM 1.5: Grafikler ve Tümünü Seç özellikleri aktiftir.")

    with st.sidebar:
        st.header("Menü")
        st.info("Sürüm 1.6 TAM")
        page = st.radio("Sayfa", ["Yeni Kayıt", "Kayıtlar", "Grafikler", "Yönetici Paneli", "Özet"], index=0)
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
    elif page == "Grafikler":
        grafikler_page()
    elif page == "Yönetici Paneli":
        yonetici_paneli_page()
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
