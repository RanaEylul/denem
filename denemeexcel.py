import os
import re
import pandas as pd
import streamlit as st
import plotly.express as px

# 📁 Kalıcı Depolama Klasörleri
CUMUL_DIR = "uploaded_cumul_files"
MONTHLY_DIR = "uploaded_monthly_files"

for d in [CUMUL_DIR, MONTHLY_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

st.set_page_config(page_title="Motor Departmanı Bütçe Takip Uygulaması", layout="wide")

# 🗓️ Kronolojik Ay Tanım Haritası
AYLAR_MAP = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
    7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
}

def ay_bilgisi_al(dosya_adi):
    match = re.search(r'M(\d{2})', dosya_adi, re.IGNORECASE)
    if match:
        ay_no = int(match.group(1))
        if 1 <= ay_no <= 12:
            return ay_no, AYLAR_MAP[ay_no]
    return 99, dosya_adi

def excel_oku_gelismis(file_source, dosya_adi):
    ay_no, ay_adi = ay_bilgisi_al(dosya_adi)
    try:
        xl = pd.ExcelFile(file_source)
    except Exception:
        return None
        
    target_sheet = None
    for s in xl.sheet_names:
        if 'cost center' in s.lower() or 'masse' in s.lower():
            target_sheet = s
            break
    if not target_sheet and xl.sheet_names:
        target_sheet = xl.sheet_names[0]
        
    if not target_sheet:
        return None

    try:
        df = pd.read_excel(file_source, sheet_name=target_sheet, skiprows=1)
        # Sütun düzenlemesi
        if df.shape[1] >= 9:
            df = df.iloc[:, 1:9]
            df.columns = ['Department', 'Cost_Center', 'Nature', 'Ref', 'Ref_Adj', 'Parity', 'Performance', 'Actual']
            df['Cost_Center'] = df['Cost_Center'].ffill()
            df['Department'] = df['Department'].ffill()
            df = df.dropna(subset=['Nature'])
            df = df[df['Cost_Center'].astype(str).str.startswith('TY') & ~df['Cost_Center'].astype(str).str.contains('Toplam') & ~df['Nature'].astype(str).str.contains('Toplam')]
            
            for col in ['Ref', 'Ref_Adj', 'Parity', 'Performance', 'Actual']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
            df['Ay_Sira'] = ay_no
            df['Ay'] = ay_adi
            return df
    except Exception:
        pass
    return None

# --- ÜST BAŞLIK ALANI (Görsel 2 ve 6'ya uygun) ---
st.markdown("### 🐎 MOTOR DEPARTMANI · TÜM RUBRİKLER · KÜMÜLATİF OCAK-EYLÜL 2026 (M09)")
st.title("Motor Departmanı Bütçe Takip Uygulaması[cite: 6]")

# --- DOSYA YÜKLEME PANELİ (Görsel 6) ---
with st.container():
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.markdown("**BÜTÇE DOSYASI**[cite: 6]")
        up_c = st.file_uploader("Kümülatif Bütçe Dosyası Yükleyin", type=["xlsx", "xls"], key="up_c_file")
        if up_c:
            with open(os.path.join(CUMUL_DIR, up_c.name), "wb") as f:
                f.write(up_c.getbuffer())
    with col_u2:
        st.markdown("**MALZEME ÇEKİŞLERİ (PHF Raporu)**[cite: 6]")
        up_phf = st.file_uploader("PHF Raporu Yükleyin", type=["xlsx", "xls"], key="up_phf_file")
        if up_phf:
            with open(os.path.join(CUMUL_DIR, up_phf.name), "wb") as f:
                f.write(up_phf.getbuffer())

saved_files = [os.path.join(CUMUL_DIR, f) for f in os.listdir(CUMUL_DIR) if f.lower().endswith(('.xlsx', '.xls'))]

if saved_files:
    all_data_list = []
    for f_path in saved_files:
        res = excel_oku_gelismis(f_path, os.path.basename(f_path))
        if res is not None and not res.empty:
            all_data_list.append(res)
            
    if all_data_list:
        df_master = pd.concat(all_data_list, ignore_index=True)
        
        # --- PARAMETRELER / FİLTRE PANELİ (Görsel 6) ---
        st.markdown("---")
        st.subheader("Parametreler")
        f_col1, f_col2, f_col3 = st.columns(3)
        
        with f_col1:
            bolumler = ["Tüm bölümler"] + sorted(df_master['Department'].dropna().unique().tolist())
            sec_bolum = st.selectbox("BÖLÜM", bolumler)
        with f_col2:
            cc_listesi = ["Tüm Motor masraf yerleri"] + sorted(df_master['Cost_Center'].dropna().unique().tolist())
            sec_cc = st.selectbox("MASRAF YERİ", cc_listesi)
        with f_col3:
            rubrikler = ["Tüm rubrikler"] + sorted(df_master['Nature'].dropna().unique().tolist())
            sec_rubrik = st.selectbox("RUBRİK", rubrikler)
            
        f_col4, f_col5 = st.columns(2)
        with f_col4:
            sec_butce_sutunu = st.selectbox("BÜTÇE SÜTUNU", ["Toplam Ref Semi Adjusted", "Reference", "Actual"])
        with f_col5:
            sec_ay = st.selectbox("KÜMÜLATİF AY", ["M09 — Ocak-Eylül", "M08 — Ocak-Ağustos", "M07 — Ocak-Temmuz"])
            
        if st.button("Filtreleri temizle"):
            st.rerun()
            
        # Filtreleme Uygula
        df_filtered = df_master.copy()
        if sec_bolum != "Tüm bölümler":
            df_filtered = df_filtered[df_filtered['Department'] == sec_bolum]
        if sec_cc != "Tüm Motor masraf yerleri":
            df_filtered = df_filtered[df_filtered['Cost_Center'] == sec_cc]
        if sec_rubrik != "Tüm rubrikler":
            df_filtered = df_filtered[df_filtered['Nature'] == sec_rubrik]

        # --- ÜST KPI ÖZET KARTLARI (Görsel 3) ---
        st.markdown("---")
        toplam_butce = df_filtered['Ref_Adj'].sum() * 1.33  # Yıllık tahmini katsayı
        harcanan = df_filtered['Performance'].sum()
        yil_sonu_kalan = toplam_butce - harcanan
        kullanim_orani = (harcanan / toplam_butce * 100) if toplam_butce > 0 else 0

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        with kpi1:
            st.metric(label="YILLIK BÜTÇE (TAHMİNİ)[cite: 5]", value=f"€ {toplam_butce:,.0f}".replace(",", "."), delta="Oca-Eylül: € 3.269.266"[cite: 5])
        with kpi2:
            st.metric(label="HARCANAN[cite: 5]", value=f"€ {harcanan:,.0f}".replace(",", "."), delta="Eylül sonu itibarıyla · dönem kullanımı %88"[cite: 5])
        with kpi3:
            st.metric(label="YIL SONUNA KALAN[cite: 5]", value=f"€ {yil_sonu_kalan:,.0f}".replace(",", "."), delta="Ekim-Aralık için"[cite: 5])
        with kpi4:
            st.metric(label="YILLIK KULLANIM[cite: 5]", value=f"%{kullanim_orani:.0f}", delta="Yılın geçen kısmı: %75"[cite: 5])

        # --- ANA DETAY TABLOSU VE İLERLEME ÇUBUKLARI (Görsel 4 ve 5) ---
        st.markdown("---")
        st.subheader("Masraf Yeri Bazlı Detaylı Bütçe ve Harcama Takibi")
        
        # Masraf yeri bazlı özet tablo oluşturma
        cc_summary = df_filtered.groupby(['Department', 'Cost_Center']).agg({
            'Ref_Adj': 'sum',
            'Performance': 'sum'
        }).reset_index()
        
        cc_summary['Donem_Kullanim'] = (cc_summary['Performance'] / cc_summary['Ref_Adj'] * 100).fillna(0)
        cc_summary['Yillik_Butce'] = cc_summary['Ref_Adj'] * 1.33
        cc_summary['Yil_Sonuna_Kalan'] = cc_summary['Yillik_Butce'] - cc_summary['Performance']
        cc_summary['Yillik_Kullanim'] = cc_summary['Donem_Kullanim'] # Örnek oran

        # Tablo Başlıkları (Görsel 4 formatında)
        header_cols = st.columns([2, 1, 1.2, 1.2, 1.2, 1, 1.2, 1.2, 1, 1.2])
        headers = ["MASRAF YERİ", "KOD", "BÜTÇE OCAK-EYLÜL (€)", "HARCANAN (€)", "DÖNEM KULLANIMI", "DÖNEM BÜTÇE AŞIMI", "YILLIK BÜTÇE (€)", "YIL SONUNA KALAN (€)", "YILLIK BÜTÇE AŞIMI", "YILLIK KULLANIM"]
        
        for col, h in zip(header_cols, headers):
            col.markdown(f"**<span style='font-size:11px;'>{h}</span>**", unsafe_allow_html=True)

        st.markdown("<hr style='margin:5px 0;'>", unsafe_allow_html=True)

        # Satırları listeleme
        for _, row in cc_summary.iterrows():
            r_cols = st.columns([2, 1, 1.2, 1.2, 1.2, 1, 1.2, 1.2, 1, 1.2])
            
            r_cols[0].markdown(f"▶ **{row['Department']}**")
            r_cols[1].markdown(f"`{row['Cost_Center']}`")
            r_cols[2].text(f"{row['Ref_Adj']:,.0f}".replace(",", "."))
            r_cols[3].text(f"{row['Performance']:,.0f}".replace(",", "."))
            
            # İlerleme çubuğu ve kullanım yüzdesi
            oran = row['Donem_Kullanim']
            renk = "#c62828" if oran > 100 else "#2e7d32"
            r_cols[4].markdown(f"<span style='color:{renk}; font-weight:bold;'>%{oran:.0f}</span>", unsafe_allow_html=True)
            
            asim = row['Performance'] - row['Ref_Adj'] if row['Performance'] > row['Ref_Adj'] else 0
            r_cols[5].text(f"{asim:,.0f}" if asim > 0 else "—")
            
            r_cols[6].text(f"{row['Yillik_Butce']:,.0f}".replace(",", "."))
            
            kalan = row['Yil_Sonuna_Kalan']
            r_cols[7].markdown(f"<span style='color: {'#c62828' if kalan < 0 else 'inherit'};'>{kalan:,.0f}</span>".replace(",", "."), unsafe_allow_html=True)
            
            r_cols[8].text("—" if asim == 0 else f"{asim:,.0f}")
            r_cols[9].markdown(f"**%{oran:.0f}**")
            st.markdown("<hr style='margin:2px 0; border-color:#f0f0f0;'>", unsafe_allow_html=True)

    else:
        st.info("💡 Yüklenen Excel dosyalarından veri okunamadı. Lütfen doğru formatta dosya yüklediğinizden emin olun.")
else:
    st.info("💡 Uygulamayı başlatmak için lütfen sol üstteki alandan Kümülatif Excel dosyanızı yükleyin.")