"""
Q-Gate Dashboard — Rzeszów
Analiza próbek + Karty kontrolne SPC (Xbar-R)
"""

import pandas as pd
import streamlit as st
import datetime
import io
import re

st.set_page_config(page_title="Q-Gate Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d5986 100%);
        border-radius: 12px; padding: 1.2rem 1.5rem;
        color: white; text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
    }
    .metric-card .label {
        font-size: 0.75rem; opacity: 0.8; margin-bottom: 4px;
        letter-spacing: 0.06em; text-transform: uppercase;
    }
    .metric-card .value { font-size: 2.1rem; font-weight: 700; line-height: 1; }
    .metric-card .sub   { font-size: 0.72rem; opacity: 0.6; margin-top: 4px; }
    header { visibility: hidden; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0; padding: 8px 20px;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)

# ─── SILNIK ANALITYCZNY ───────────────────────────────────────────────────────

KOLUMNY_PROBEK = [
    'Wysokość zaciskania - zrywy',
    'Wysokość zaciskania -  kontrola przekroju na mikrografie',
    'Wysokość zaciskania - próbka poglądowa',
]
LINIE_DO_IGNOROWANIA = {'`', '', 'NAN'}


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
    if 'HPC' in t and 'UNCOOL' in t:
        return 'HPC UNCOOLED'
    if 'HPC' in t and '1' in t:
        return 'HPC 1.0'
    if 'HPC' in t and '2' in t:
        return 'HPC 2.0'
    return t


@st.cache_data(show_spinner='Wczytuję dane...')
def wczytaj_dane(plik_excel):
    df = pd.read_excel(plik_excel, sheet_name='Qgate 2026', header=2, engine='openpyxl').copy()

    for col in ['Numer zlecenia ', 'NR LINII ', 'Numer kontaktu', 'Typ pojedyńczego przewodu']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace(['nan', 'None', 'nan.0'], '')

    df['NR LINII '] = df['NR LINII '].apply(normalizuj_linie)
    df['Numer zlecenia '] = df['Numer zlecenia '].replace('', pd.NA).ffill()
    df['NR LINII '] = df['NR LINII '].replace('', pd.NA).ffill()
    df = df[~df['NR LINII '].isin(LINIE_DO_IGNOROWANIA)]

    df['Data i czas startu zlecenia'] = pd.to_datetime(
        df['Data i czas startu zlecenia'], errors='coerce').ffill()
    df['Data_kalendarzowa'] = df['Data i czas startu zlecenia'].dt.date

    for col in KOLUMNY_PROBEK:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df['Wiersz_ma_probke'] = df[KOLUMNY_PROBEK].notna().any(axis=1)

    # Liczymy TYLKO wiersze z rzeczywistym pomiarem wysokości zaciskania.
    # Kontakty bez wpisanych wartości (np. PP zakrępowany z rezystorem)
    # nie trafiają do raportu — nie mamy ich pomiaru więc ich nie scrappujemy.
    zlecenia_status = df.groupby(
        ['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa']
    )['Wiersz_ma_probke'].any().reset_index()
    zestawy_aktywne = zlecenia_status[zlecenia_status['Wiersz_ma_probke']]

    df_zestawy = pd.merge(
        df, zestawy_aktywne[['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa']],
        on=['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa'], how='inner',
    )
    # Tylko wiersze z pomiarem — żadnych wyjątków dla PP/DC bez danych
    # Klucz deduplikacji BEZ Data_kalendarzowa — ten sam kontakt mierzony
    # na początku i końcu produkcji (różne dni) to nadal jedna fizyczna próbka.
    df_unikalne = df_zestawy[df_zestawy['Wiersz_ma_probke']].drop_duplicates(
        subset=['Numer zlecenia ', 'NR LINII ',
                'Numer kontaktu', 'Typ pojedyńczego przewodu']
    ).copy()
    df_unikalne['Sztuki_fizyczne'] = 3

    zlecenia_wynik = df_unikalne.groupby(
        ['Numer zlecenia ', 'NR LINII ']
    ).agg(
        Suma_kontaktow=('Numer kontaktu', 'size'),
        Data_max=('Data i czas startu zlecenia', 'max'),
    ).reset_index().rename(columns={'Data_max': 'Data i czas startu zlecenia'})

    # Data_kalendarzowa wyznaczona z daty pierwszego pomiaru w zleceniu
    zlecenia_wynik['Data_kalendarzowa'] = zlecenia_wynik['Data i czas startu zlecenia'].dt.date
    zlecenia_wynik['Ilosc_zestawow'] = 1
    zlecenia_wynik['Ilosc_fizycznych_probek'] = zlecenia_wynik['Suma_kontaktow'] * 3

    return zlecenia_wynik, df_unikalne


def wygeneruj_excel(df_linia, df_sap, df_trend=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_linia.to_excel(writer, index=False, sheet_name='Podsumowanie Linii')
        df_sap.to_excel(writer, index=False, sheet_name='Ściągawka SAP')
        if df_trend is not None:
            df_trend.to_excel(writer, index=False, sheet_name='Trend Dzienny')
    return output.getvalue()


# ─── SIDEBAR ─────────────────────────────────────────────────────────────────

st.title('📊 Q-Gate Dashboard — Rzeszów')

with st.sidebar:
    st.header('📁 Wczytaj dane')
    plik = st.file_uploader('Wgraj plik Q-Gate (xlsx, xlsm):', type=['xlsx', 'xlsm'])

if not plik:
    st.info('👈 Wgraj plik produkcyjny Q-Gate żeby rozpocząć.')
    st.markdown("""
    **Obsługiwane formaty:** `.xlsx`, `.xlsm`  
    **Wymagany arkusz:** `Qgate 2026`  
    **Nagłówek:** wiersz 3
    """)
    st.stop()

# ─── WCZYTANIE ───────────────────────────────────────────────────────────────

try:
    df_zlecenia, df_kontakty = wczytaj_dane(plik)
except Exception as e:
    st.error(f'❌ Błąd podczas wczytywania: {e}')
    st.stop()

if df_zlecenia.empty:
    st.warning('⚠️ Plik nie zawiera danych spełniających kryteria.')
    st.stop()

# ─── ZAKŁADKI ────────────────────────────────────────────────────────────────

tab_raport, tab_spc = st.tabs(['📋 Raport produkcji', '📈 Karty SPC'])

# ══════════════════════════════════════════════════════════════════════════════
# ZAKŁADKA 1 — RAPORT PRODUKCJI
# ══════════════════════════════════════════════════════════════════════════════

with tab_raport:
    with st.sidebar:
        st.markdown('---')
        tryb = st.radio('Rodzaj raportu:', ['Dzienny / Zmianowy', 'Miesięczny', 'Zakres dat'])
        dostepne_linie = sorted(df_zlecenia['NR LINII '].dropna().unique())
        wybrane_linie = st.multiselect('Filtruj linie (puste = wszystkie):', dostepne_linie)

        if tryb == 'Dzienny / Zmianowy':
            min_d = df_zlecenia['Data i czas startu zlecenia'].min().date()
            max_d = df_zlecenia['Data i czas startu zlecenia'].max().date()
            wybrana_d = st.date_input('Dzień:', value=max_d, min_value=min_d, max_value=max_d)
            c1, c2 = st.columns(2)
            od_g = c1.time_input('Od:', value=datetime.time(6, 0))
            do_g = c2.time_input('Do:', value=datetime.time(14, 0))
            s_dt = pd.to_datetime(f'{wybrana_d} {od_g}')
            e_dt = pd.to_datetime(f'{wybrana_d} {do_g}')
            if do_g <= od_g:
                e_dt += pd.Timedelta(days=1)
            df_f = df_zlecenia[
                (df_zlecenia['Data i czas startu zlecenia'] >= s_dt) &
                (df_zlecenia['Data i czas startu zlecenia'] <= e_dt)
            ]
            df_k_f = df_kontakty[
                (df_kontakty['Data i czas startu zlecenia'] >= s_dt) &
                (df_kontakty['Data i czas startu zlecenia'] <= e_dt)
            ]
            tytul = f"Zmiana: {s_dt.strftime('%d.%m.%Y %H:%M')} – {e_dt.strftime('%H:%M')}"

        elif tryb == 'Miesięczny':
            df_zlecenia['YM'] = df_zlecenia['Data i czas startu zlecenia'].dt.strftime('%Y-%m')
            df_kontakty['YM'] = df_kontakty['Data i czas startu zlecenia'].dt.strftime('%Y-%m')
            wybrany_m = st.selectbox('Miesiąc:', sorted(df_zlecenia['YM'].unique(), reverse=True))
            df_f = df_zlecenia[df_zlecenia['YM'] == wybrany_m]
            df_k_f = df_kontakty[df_kontakty['YM'] == wybrany_m]
            tytul = f'Raport miesięczny: {wybrany_m}'

        else:
            min_d = df_zlecenia['Data i czas startu zlecenia'].min().date()
            max_d = df_zlecenia['Data i czas startu zlecenia'].max().date()
            dr = st.date_input('Zakres:', value=(min_d, max_d), min_value=min_d, max_value=max_d)
            if len(dr) == 2:
                s_dt = pd.to_datetime(dr[0])
                e_dt = pd.to_datetime(dr[1]) + pd.Timedelta(days=1)
                df_f = df_zlecenia[
                    (df_zlecenia['Data i czas startu zlecenia'] >= s_dt) &
                    (df_zlecenia['Data i czas startu zlecenia'] < e_dt)
                ]
                df_k_f = df_kontakty[
                    (df_kontakty['Data i czas startu zlecenia'] >= s_dt) &
                    (df_kontakty['Data i czas startu zlecenia'] < e_dt)
                ]
                tytul = f"Zakres: {dr[0].strftime('%d.%m.%Y')} – {dr[1].strftime('%d.%m.%Y')}"
            else:
                st.warning('Wybierz zakres dat.')
                st.stop()

    if wybrane_linie:
        df_f = df_f[df_f['NR LINII '].isin(wybrane_linie)]
        df_k_f = df_k_f[df_k_f['NR LINII '].isin(wybrane_linie)]

    st.subheader(tytul)

    if df_f.empty:
        st.warning('⚠️ Brak danych dla wybranych filtrów.')
    else:
        c1, c2, c3, c4 = st.columns(4)
        for col, lbl, val, sub in [
            (c1, '🔥 Próbki do SAP', f"{int(df_f['Ilosc_fizycznych_probek'].sum()):,} szt.", 'do wprowadzenia'),
            (c2, '📦 Zestawy', f"{int(df_f['Ilosc_zestawow'].sum()):,} szt.", 'zrealizowanych'),
            (c3, '🔌 Unikalne PN', f"{int(df_f['Suma_kontaktow'].sum()):,} szt.", 'part numberów'),
            (c4, '🏭 Aktywne linie', str(df_f['NR LINII '].nunique()), 'linii produkcyjnych'),
        ]:
            with col:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="label">{lbl}</div>
                    <div class="value">{val}</div>
                    <div class="sub">{sub}</div>
                </div>""", unsafe_allow_html=True)

        st.markdown('<br>', unsafe_allow_html=True)
        st.divider()

        pod_linia = df_f.groupby('NR LINII ').agg(
            Kontakty=('Suma_kontaktow', 'sum'),
            Zestawy=('Ilosc_zestawow', 'sum'),
            Sztuki_SAP=('Ilosc_fizycznych_probek', 'sum'),
        ).reset_index().rename(columns={'NR LINII ': 'Linia'}).sort_values('Sztuki_SAP', ascending=False)
        pod_linia['Udział %'] = (pod_linia['Sztuki_SAP'] / pod_linia['Sztuki_SAP'].sum() * 100).round(1)

        ct, cw = st.columns([1, 1.3])
        with ct:
            st.markdown('**📋 Dane per linia**')
            st.dataframe(
                pod_linia.style.background_gradient(subset=['Sztuki_SAP'], cmap='Blues'),
                use_container_width=True, hide_index=True,
            )
        with cw:
            st.markdown('**📊 Próbki SAP per linia**')
            st.bar_chart(pod_linia.set_index('Linia')['Sztuki_SAP'], color='#2d5986')

        if tryb in ('Miesięczny', 'Zakres dat'):
            st.divider()
            trend = df_f.groupby('Data_kalendarzowa')['Ilosc_fizycznych_probek'].sum().reset_index()
            trend['Dzień'] = pd.to_datetime(trend['Data_kalendarzowa']).dt.strftime('%d.%m')
            ct2, cs = st.columns([2, 1])
            with ct2:
                st.markdown('**📈 Trend dzienny (sztuki SAP)**')
                st.bar_chart(trend.set_index('Dzień')['Ilosc_fizycznych_probek'], color='#1e7a4e')
            with cs:
                st.markdown('**📉 Statystyki dzienne**')
                dv = trend['Ilosc_fizycznych_probek']
                st.metric('Średnio / dzień', f'{dv.mean():.0f} szt.')
                st.metric('Maks. dzień', f'{dv.max():.0f} szt.')
                st.metric('Min. dzień', f'{dv.min():.0f} szt.')
                st.metric('Dni roboczych', str(len(trend)))

        st.divider()
        st.markdown('### 🗂️ Ściągawka do SAP')

        sap_tab = df_k_f.groupby('Numer kontaktu')['Sztuki_fizyczne'].sum().reset_index()
        sap_tab = sap_tab.rename(columns={'Numer kontaktu': 'PN', 'Sztuki_fizyczne': 'Sztuki'})
        linia_info = df_k_f.groupby('Numer kontaktu')['NR LINII '].agg(
            lambda x: ', '.join(sorted(set(x)))).reset_index()
        linia_info.columns = ['PN', 'Linie']
        sap_tab = sap_tab.merge(linia_info, on='PN').sort_values('Sztuki', ascending=False).reset_index(drop=True)

        cs1, cs2 = st.columns([2, 1])
        with cs1:
            st.dataframe(sap_tab, use_container_width=True, hide_index=True, height=280)
        with cs2:
            st.metric('Unikalnych PN', str(len(sap_tab)))
            st.metric('Łącznie sztuk', f"{sap_tab['Sztuki'].sum():,}")
            with st.expander('📄 Szczegółowe zlecenia', expanded=False):
                st.dataframe(
                    df_f[['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa',
                           'Suma_kontaktow', 'Ilosc_fizycznych_probek']]
                    .rename(columns={
                        'Numer zlecenia ': 'Zlecenie', 'NR LINII ': 'Linia',
                        'Data_kalendarzowa': 'Data', 'Suma_kontaktow': 'PN',
                        'Ilosc_fizycznych_probek': 'Szt. SAP',
                    }).reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                )

        st.divider()
        trend_exp = None
        if tryb in ('Miesięczny', 'Zakres dat'):
            trend_exp = df_f.groupby('Data_kalendarzowa')['Ilosc_fizycznych_probek'].sum().reset_index()
            trend_exp.columns = ['Data', 'Sztuki_SAP']

        cd1, cd2, _ = st.columns([1, 1, 2])
        with cd1:
            st.download_button(
                '📥 Pobierz Excel',
                data=wygeneruj_excel(pod_linia, sap_tab, trend_exp),
                file_name=f"QGate_{datetime.date.today()}.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
        with cd2:
            st.download_button(
                '📄 Pobierz SAP CSV',
                data=sap_tab.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig'),
                file_name=f"SAP_{datetime.date.today()}.csv",
                mime='text/csv',
            )

# ══════════════════════════════════════════════════════════════════════════════
# ZAKŁADKA 2 — SPC
# ══════════════════════════════════════════════════════════════════════════════

with tab_spc:
    from spc_qgate import pokaz_spc
    pokaz_spc(plik)