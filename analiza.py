import pandas as pd
import streamlit as st
import datetime
import io
import re

st.set_page_config(page_title="Analizator Q-Gate", layout="wide", page_icon="📊")

st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d5986 100%);
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        color: white;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
    }
    .metric-card .label { font-size: 0.85rem; opacity: 0.8; margin-bottom: 0.3rem; letter-spacing: 0.05em; text-transform: uppercase; }
    .metric-card .value { font-size: 2.2rem; font-weight: 700; line-height: 1; }
    .metric-card .sub { font-size: 0.75rem; opacity: 0.6; margin-top: 0.2rem; }
    div[data-testid="stMetricValue"] { font-size: 1.8rem !important; }
    .stAlert { border-radius: 8px; }
    header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─── SILNIK ANALITYCZNY ───────────────────────────────────────────────────────

KOLUMNY_PROBEK = [
    'Wysokość zaciskania - zrywy',
    'Wysokość zaciskania -  kontrola przekroju na mikrografie',
    'Wysokość zaciskania - próbka poglądowa',
]

LINIE_DO_IGNOROWANIA = {'`', '', 'NAN'}  # backtick to błąd danych


def normalizuj_linie(tekst):
    if pd.isna(tekst) or str(tekst).lower() in ['nan', 'none', '']:
        return ''
    t = str(tekst).upper().strip()
    t = re.sub(r'[\.\-/_]', ' ', t)
    t = ' '.join(t.split())
    if 'UNI INLET' in t or t == 'INLET':
        return 'INLET'
    if 'CCSD' in t.replace(' ', '') or t == 'CCS D':
        return 'CCS D'
    return t


@st.cache_data(show_spinner="Wczytuję i analizuję dane...")
def wczytaj_dane(plik_excel):
    df = pd.read_excel(plik_excel, sheet_name='Qgate 2026', header=2, engine='openpyxl').copy()

    for col in ['Numer zlecenia ', 'NR LINII ', 'Numer kontaktu', 'Typ pojedyńczego przewodu']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace(['nan', 'None', 'nan.0'], '')

    df['NR LINII '] = df['NR LINII '].apply(normalizuj_linie)
    df['Numer zlecenia '] = df['Numer zlecenia '].replace('', pd.NA).ffill()
    df['NR LINII '] = df['NR LINII '].replace('', pd.NA).ffill()

    # Usuń błędne linie (np. backtick z błędów OCR)
    df = df[~df['NR LINII '].isin(LINIE_DO_IGNOROWANIA)]

    df['Data i czas startu zlecenia'] = pd.to_datetime(df['Data i czas startu zlecenia'], errors='coerce').ffill()
    df['Data_kalendarzowa'] = df['Data i czas startu zlecenia'].dt.date

    for col in KOLUMNY_PROBEK:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        else:
            df[col] = float('nan')

    df['Wiersz_ma_probke'] = df[KOLUMNY_PROBEK].notna().any(axis=1)
    df['Czy_specjalny'] = df['Typ pojedyńczego przewodu'].str.contains(r'PP|DC\+|DC-', regex=True, na=False)

    zlecenia_status = df.groupby(['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa'])['Wiersz_ma_probke'].any().reset_index()
    zestawy_aktywne = zlecenia_status[zlecenia_status['Wiersz_ma_probke']]

    df_zestawy = pd.merge(
        df, zestawy_aktywne[['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa']],
        on=['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa'], how='inner'
    )

    df_do_zliczenia = df_zestawy[df_zestawy['Wiersz_ma_probke'] | df_zestawy['Czy_specjalny']]

    df_unikalne = df_do_zliczenia.drop_duplicates(
        subset=['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa', 'Numer kontaktu', 'Typ pojedyńczego przewodu']
    ).copy()
    df_unikalne['Sztuki_fizyczne'] = 3

    zlecenia_wynik = df_unikalne.groupby(['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa']).agg(
        Suma_kontaktow=('Numer kontaktu', 'size'),
        Data_max=('Data i czas startu zlecenia', 'max')
    ).reset_index().rename(columns={'Data_max': 'Data i czas startu zlecenia'})

    zlecenia_wynik['Ilosc_zestawow'] = 1
    zlecenia_wynik['Ilosc_fizycznych_probek'] = zlecenia_wynik['Suma_kontaktow'] * 3

    return zlecenia_wynik, df_unikalne


# ─── GENERATOR EXCEL ─────────────────────────────────────────────────────────

def wygeneruj_excel(df_linia, df_sap, df_trend=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_linia.to_excel(writer, index=False, sheet_name='Podsumowanie Linii')
        df_sap.to_excel(writer, index=False, sheet_name='Ściągawka SAP')
        if df_trend is not None:
            df_trend.to_excel(writer, index=False, sheet_name='Trend Dzienny')
    return output.getvalue()


# ─── INTERFEJS ───────────────────────────────────────────────────────────────

st.title("📊 Dashboard analizy próbek Q-Gate")

with st.sidebar:
    st.header("📁 Wczytaj dane")
    plik = st.file_uploader("Wgraj plik Q-Gate (xlsx, xlsm):", type=['xlsx', 'xlsm'])

    if plik:
        st.markdown("---")
        tryb = st.radio("Rodzaj raportu:", ["Dzienny / Zmianowy", "Miesięczny", "Zakres dat"])

if not plik:
    col1, col2 = st.columns([1, 2])
    with col1:
        st.info("👈 Wgraj plik produkcyjny Q-Gate, aby rozpocząć analizę.")
        st.markdown("""
        **Obsługiwane formaty:** `.xlsx`, `.xlsm`  
        **Wymagany arkusz:** `Qgate 2026`  
        **Nagłówek danych:** wiersz 3
        """)
    st.stop()

# ─── WCZYTANIE I ANALIZA ─────────────────────────────────────────────────────

try:
    df_zlecenia, df_kontakty = wczytaj_dane(plik)
except Exception as e:
    st.error(f"❌ Błąd podczas wczytywania pliku: {e}")
    st.info("Sprawdź czy plik zawiera arkusz 'Qgate 2026' z nagłówkiem w wierszu 3.")
    st.stop()

if df_zlecenia.empty:
    st.warning("⚠️ Plik nie zawiera danych spełniających kryteria (brak wierszy z pomiarami).")
    st.stop()

# ─── FILTRY SIDEBAR ──────────────────────────────────────────────────────────

with st.sidebar:
    # Filtr linii
    dostepne_linie = sorted(df_zlecenia['NR LINII '].dropna().unique())
    wybrane_linie = st.multiselect("Filtruj linie (puste = wszystkie):", dostepne_linie)

    if tryb == "Dzienny / Zmianowy":
        min_data = df_zlecenia['Data i czas startu zlecenia'].min().date()
        max_data = df_zlecenia['Data i czas startu zlecenia'].max().date()
        wybrana_d = st.date_input("Wybierz dzień", value=max_data, min_value=min_data, max_value=max_data)
        col_od, col_do = st.columns(2)
        with col_od:
            od_g = st.time_input("Od:", value=datetime.time(6, 0))
        with col_do:
            do_g = st.time_input("Do:", value=datetime.time(14, 0))

        s_dt = pd.to_datetime(f"{wybrana_d} {od_g}")
        e_dt = pd.to_datetime(f"{wybrana_d} {do_g}")
        if do_g <= od_g:
            e_dt += pd.Timedelta(days=1)

        df_f = df_zlecenia[(df_zlecenia['Data i czas startu zlecenia'] >= s_dt) &
                           (df_zlecenia['Data i czas startu zlecenia'] <= e_dt)]
        df_k_f = df_kontakty[(df_kontakty['Data i czas startu zlecenia'] >= s_dt) &
                              (df_kontakty['Data i czas startu zlecenia'] <= e_dt)]
        tytul = f"Zmiana: {s_dt.strftime('%d.%m.%Y %H:%M')} – {e_dt.strftime('%H:%M')}"

    elif tryb == "Miesięczny":
        df_zlecenia['YM'] = df_zlecenia['Data i czas startu zlecenia'].dt.strftime('%Y-%m')
        df_kontakty['YM'] = df_kontakty['Data i czas startu zlecenia'].dt.strftime('%Y-%m')
        dostepne_miesiace = sorted(df_zlecenia['YM'].unique(), reverse=True)
        wybrany_m = st.selectbox("Wybierz miesiąc:", dostepne_miesiace)
        df_f = df_zlecenia[df_zlecenia['YM'] == wybrany_m]
        df_k_f = df_kontakty[df_kontakty['YM'] == wybrany_m]
        tytul = f"Raport miesięczny: {wybrany_m}"

    else:  # Zakres dat
        min_data = df_zlecenia['Data i czas startu zlecenia'].min().date()
        max_data = df_zlecenia['Data i czas startu zlecenia'].max().date()
        date_range = st.date_input("Wybierz zakres dat:", value=(min_data, max_data),
                                   min_value=min_data, max_value=max_data)
        if len(date_range) == 2:
            s_dt = pd.to_datetime(date_range[0])
            e_dt = pd.to_datetime(date_range[1]) + pd.Timedelta(days=1)
            df_f = df_zlecenia[(df_zlecenia['Data i czas startu zlecenia'] >= s_dt) &
                               (df_zlecenia['Data i czas startu zlecenia'] < e_dt)]
            df_k_f = df_kontakty[(df_kontakty['Data i czas startu zlecenia'] >= s_dt) &
                                  (df_kontakty['Data i czas startu zlecenia'] < e_dt)]
            tytul = f"Zakres: {date_range[0].strftime('%d.%m.%Y')} – {date_range[1].strftime('%d.%m.%Y')}"
        else:
            st.warning("Wybierz zakres dat.")
            st.stop()

# Filtr linii
if wybrane_linie:
    df_f = df_f[df_f['NR LINII '].isin(wybrane_linie)]
    df_k_f = df_k_f[df_k_f['NR LINII '].isin(wybrane_linie)]

# ─── GŁÓWNA TREŚĆ ────────────────────────────────────────────────────────────

st.subheader(tytul)

if df_f.empty:
    st.warning("⚠️ Brak danych dla wybranych filtrów. Zmień zakres dat lub linie.")
    st.stop()

# Metryki
c1, c2, c3, c4 = st.columns(4)

fizyczne = int(df_f['Ilosc_fizycznych_probek'].sum())
zestawy = int(df_f['Ilosc_zestawow'].sum())
kontakty = int(df_f['Suma_kontaktow'].sum())
aktywne_linie = df_f['NR LINII '].nunique()

for col, label, val, sub in [
    (c1, "🔥 Próbki do SAP", f"{fizyczne:,} szt.", "do wprowadzenia"),
    (c2, "📦 Zestawy", f"{zestawy:,} szt.", "zrealizowanych"),
    (c3, "🔌 Kontakty", f"{kontakty:,} szt.", "zbrakowanych"),
    (c4, "🏭 Aktywne linie", str(aktywne_linie), "linii produkcyjnych"),
]:
    with col:
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">{label}</div>
            <div class="value">{val}</div>
            <div class="sub">{sub}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.divider()

# Tabela i wykres per linia
col_t, col_w = st.columns([1, 1.3])

pod_linia = df_f.groupby('NR LINII ').agg(
    Kontakty=('Suma_kontaktow', 'sum'),
    Zestawy=('Ilosc_zestawow', 'sum'),
    Sztuki_SAP=('Ilosc_fizycznych_probek', 'sum')
).reset_index().rename(columns={'NR LINII ': 'Linia'}).sort_values('Sztuki_SAP', ascending=False)

# Oblicz % udziału
pod_linia['Udział %'] = (pod_linia['Sztuki_SAP'] / pod_linia['Sztuki_SAP'].sum() * 100).round(1)

with col_t:
    st.markdown("**📋 Dane per linia**")
    st.dataframe(
        pod_linia.style.background_gradient(subset=['Sztuki_SAP'], cmap='Blues'),
        use_container_width=True, hide_index=True
    )

with col_w:
    st.markdown("**📊 Próbki SAP per linia**")
    chart_data = pod_linia.set_index('Linia')['Sztuki_SAP']
    st.bar_chart(chart_data, color='#2d5986')

# Trend dzienny (miesięczny i zakres dat)
if tryb in ("Miesięczny", "Zakres dat"):
    st.divider()
    trend = df_f.groupby('Data_kalendarzowa')['Ilosc_fizycznych_probek'].sum().reset_index()
    trend['Dzień'] = pd.to_datetime(trend['Data_kalendarzowa']).dt.strftime('%d.%m')

    col_trend, col_stat = st.columns([2, 1])
    with col_trend:
        st.markdown("**📈 Trend dzienny (sztuki SAP)**")
        st.bar_chart(trend.set_index('Dzień')['Ilosc_fizycznych_probek'], color='#1e7a4e')

    with col_stat:
        st.markdown("**📉 Statystyki dzienne**")
        daily_vals = trend['Ilosc_fizycznych_probek']
        st.metric("Średnio / dzień", f"{daily_vals.mean():.0f} szt.")
        st.metric("Maks. dzień", f"{daily_vals.max():.0f} szt.")
        st.metric("Min. dzień", f"{daily_vals.min():.0f} szt.")
        st.metric("Dni roboczych", str(len(trend)))

# Ściągawka SAP
st.divider()
st.markdown("### 🗂️ Ściągawka do SAP (Numery PN)")

sap_tab = df_k_f.groupby('Numer kontaktu')['Sztuki_fizyczne'].sum().reset_index()
sap_tab = sap_tab.rename(columns={'Numer kontaktu': 'Numer kontaktu (PN)', 'Sztuki_fizyczne': 'Sztuki'})
sap_tab = sap_tab.sort_values('Sztuki', ascending=False).reset_index(drop=True)

# Łącz z informacją o linii
linia_info = df_k_f.groupby('Numer kontaktu')['NR LINII '].agg(lambda x: ', '.join(sorted(set(x)))).reset_index()
linia_info.columns = ['Numer kontaktu (PN)', 'Linie']
sap_tab = sap_tab.merge(linia_info, on='Numer kontaktu (PN)', how='left')

col_sap, col_sum = st.columns([2, 1])
with col_sap:
    st.dataframe(sap_tab, use_container_width=True, hide_index=True, height=300)
with col_sum:
    st.markdown("**Podsumowanie SAP**")
    st.metric("Unikalnych PN", str(len(sap_tab)))
    st.metric("Łącznie sztuk", f"{sap_tab['Sztuki'].sum():,}")

    # Szczegóły zlecenia
    with st.expander("📄 Szczegółowe zlecenia", expanded=False):
        st.dataframe(
            df_f[['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa', 'Suma_kontaktow', 'Ilosc_fizycznych_probek']]
            .rename(columns={
                'Numer zlecenia ': 'Zlecenie', 'NR LINII ': 'Linia',
                'Data_kalendarzowa': 'Data', 'Suma_kontaktow': 'Kontakty',
                'Ilosc_fizycznych_probek': 'Sztuki SAP'
            }).reset_index(drop=True),
            use_container_width=True, hide_index=True
        )

# Pobieranie
st.divider()
trend_exp = None
if tryb in ("Miesięczny", "Zakres dat"):
    trend_exp = df_f.groupby('Data_kalendarzowa')['Ilosc_fizycznych_probek'].sum().reset_index()
    trend_exp.columns = ['Data', 'Sztuki_SAP']

col_dl1, col_dl2, _ = st.columns([1, 1, 2])
with col_dl1:
    st.download_button(
        "📥 Pobierz raport Excel",
        data=wygeneruj_excel(pod_linia, sap_tab, trend_exp),
        file_name=f"Raport_QGate_{tytul.replace(':', '').replace(' ', '_')[:50]}_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with col_dl2:
    csv_data = sap_tab.to_csv(index=False, sep=';', encoding='utf-8-sig')
    st.download_button(
        "📄 Pobierz SAP jako CSV",
        data=csv_data.encode('utf-8-sig'),
        file_name=f"SAP_{datetime.date.today()}.csv",
        mime="text/csv"
    )
