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
    /* ── Metryki ── */
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

    /* ── Ukryj domyślny header Streamlit ── */
    header { visibility: hidden; }

    /* ── Zakładki ── */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0; padding: 8px 20px;
        font-weight: 500;
    }

    /* ── Dark mode dla sekcji SPC ── */
    /* Plotly wykresy — usuń białe obramowanie */
    .js-plotly-plot .plotly .main-svg {
        border-radius: 10px;
    }
    /* Tło legendy wykresu SPC pasuje do dark bg */
    [data-testid="stPlotlyChart"] {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 24px rgba(0,0,0,0.4);
    }

    /* ── SIDEBAR: zawsze widoczny, bez przycisku collapse ── */
    /* Ukryj przycisk strzałki zwijania */
    button[kind="header"],
    [data-testid="collapsedControl"],
    [data-testid="stSidebarCollapseButton"] {
        display: none !important;
    }
    /* Zawsze pokazuj sidebar (nawet gdy Streamlit chce go ukryć) */
    [data-testid="stSidebar"] {
        transform: none !important;
        min-width: 21rem !important;
        max-width: 21rem !important;
    }
    /* Usuń margines który pojawia się po "zwinięciu" */
    [data-testid="stSidebar"][aria-expanded="false"] {
        margin-left: 0 !important;
        visibility: visible !important;
    }
    /* Główna treść — zawsze z marginesem na sidebar */
    .main .block-container {
        padding-left: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ─── SILNIK ANALITYCZNY ───────────────────────────────────────────────────────

LINIE_DO_IGNOROWANIA = {'`', '', 'NAN'}


def _znajdz_kolumne(df, *fragmenty, wymagana=True):
    """Znajduje kolumnę zawierającą wszystkie podane fragmenty (bez względu
    na spacje/wielkość liter). Odporne na różnice między wersjami pliku."""
    def uprość(s):
        return re.sub(r'\s+', ' ', str(s).lower().strip())
    frag_low = [uprość(f) for f in fragmenty]
    for col in df.columns:
        cl = uprość(col)
        if all(f in cl for f in frag_low):
            return col
    if wymagana:
        raise KeyError(f"Nie znaleziono kolumny z fragmentami: {fragmenty}")
    return None


def _wczytaj_arkusz(plik_excel, sheet_name):
    """Wczytuje arkusz auto-wykrywając wiersz nagłówka (3 lub 2).
    W nowszych plikach Qgate 2026 nagłówek przesunął się do wiersza 3."""
    for hdr in (3, 2):
        df = pd.read_excel(plik_excel, sheet_name=sheet_name,
                           header=hdr, engine='openpyxl').copy()
        cols_str = ' '.join(str(c) for c in df.columns[:12])
        if 'Numer zlecenia' in cols_str and 'Kolumna1' not in cols_str:
            return df
    return pd.read_excel(plik_excel, sheet_name=sheet_name,
                         header=2, engine='openpyxl').copy()


def normalizuj_linie(tekst):
    if pd.isna(tekst) or str(tekst).lower() in ['nan', 'none', '']:
        return ''
    t = str(tekst).upper().strip()
    t = re.sub(r'[\.\-/_]', ' ', t)
    t = ' '.join(t.split())
    if 'NACS' in t and 'DC' in t:
        return 'NACS DC'
    if t == 'NAC' or t == 'NACS':
        return 'NACS'
    if 'UNIINLET' in t.replace(' ', '') or t == 'INLET':
        return 'INLET'
    if 'CCSD' in t.replace(' ', '') or t == 'CCS D':
        return 'CCS D'
    if 'HPC' in t and 'UNCOOL' in t:
        return 'HPC UNCOOLED'
    if 'HPC' in t and '1' in t:
        return 'HPC 1.0'
    if 'HPC' in t and '2' in t:
        return 'HPC 2.0'
    if 'PREASSY' in t and 'GEN 3' in t:
        return 'PREASSY GEN 3'
    if 'PREASSY' in t and ('GEN 2' in t or 'GEN2' in t):
        return 'PREASSY GEN 2'
    return t


@st.cache_data(show_spinner='Wczytuję dane...')
def wczytaj_dane(plik_excel):
    df = _wczytaj_arkusz(plik_excel, 'Qgate 2026')

    c_zlec    = _znajdz_kolumne(df, 'numer zlecenia')
    c_linia   = _znajdz_kolumne(df, 'nr linii')
    c_kontakt = _znajdz_kolumne(df, 'numer kontaktu')
    c_typ     = _znajdz_kolumne(df, 'typ pojedy')
    c_data    = _znajdz_kolumne(df, 'data i czas startu')
    c_zryw    = _znajdz_kolumne(df, 'zaciskania', 'zrywy')
    c_mikro   = _znajdz_kolumne(df, 'zaciskania', 'mikrografie')
    c_poglad  = _znajdz_kolumne(df, 'zaciskania', 'poglądowa')

    for col in [c_zlec, c_linia, c_kontakt, c_typ]:
        df[col] = df[col].astype(str).str.strip().replace(['nan', 'None', 'nan.0'], '')

    df['Linia'] = df[c_linia].apply(normalizuj_linie)
    df['Zlecenie'] = df[c_zlec].replace('', pd.NA).ffill()
    df['Linia'] = df['Linia'].replace('', pd.NA).ffill()
    df = df[~df['Linia'].isin(LINIE_DO_IGNOROWANIA)]

    df['Data i czas startu zlecenia'] = pd.to_datetime(df[c_data], errors='coerce').ffill()
    df['Data_kalendarzowa'] = df['Data i czas startu zlecenia'].dt.date

    df['_zryw']  = pd.to_numeric(df[c_zryw], errors='coerce')
    df['_mikro'] = pd.to_numeric(df[c_mikro], errors='coerce')
    df['_pogl']  = pd.to_numeric(df[c_poglad], errors='coerce')
    df['Kontakt'] = df[c_kontakt].astype(str).str.strip()
    df['Typ'] = df[c_typ].astype(str).str.strip()

    df['Wiersz_ma_probke'] = df[['_zryw', '_mikro', '_pogl']].notna().any(axis=1)

    # Liczymy TYLKO wiersze z rzeczywistym pomiarem wysokości zaciskania.
    zlecenia_status = df.groupby(
        ['Zlecenie', 'Linia', 'Data_kalendarzowa']
    )['Wiersz_ma_probke'].any().reset_index()
    zestawy_aktywne = zlecenia_status[zlecenia_status['Wiersz_ma_probke']]

    df_zestawy = pd.merge(
        df, zestawy_aktywne[['Zlecenie', 'Linia', 'Data_kalendarzowa']],
        on=['Zlecenie', 'Linia', 'Data_kalendarzowa'], how='inner',
    )
    # Klucz deduplikacji BEZ Data_kalendarzowa — ten sam kontakt mierzony
    # na początku i końcu produkcji (różne dni) to nadal jedna próbka.
    df_unikalne = df_zestawy[df_zestawy['Wiersz_ma_probke']].drop_duplicates(
        subset=['Zlecenie', 'Linia', 'Kontakt', 'Typ']
    ).copy()
    df_unikalne['Sztuki_fizyczne'] = 3
    # Aliasy dla zgodności z resztą kodu
    df_unikalne['Numer kontaktu'] = df_unikalne['Kontakt']
    df_unikalne['NR LINII '] = df_unikalne['Linia']

    zlecenia_wynik = df_unikalne.groupby(['Zlecenie', 'Linia']).agg(
        Suma_kontaktow=('Kontakt', 'size'),
        Data_max=('Data i czas startu zlecenia', 'max'),
    ).reset_index().rename(columns={
        'Data_max': 'Data i czas startu zlecenia',
        'Linia': 'NR LINII ',
    })

    zlecenia_wynik['Data_kalendarzowa'] = zlecenia_wynik['Data i czas startu zlecenia'].dt.date
    zlecenia_wynik['Ilosc_zestawow'] = 1
    zlecenia_wynik['Ilosc_fizycznych_probek'] = zlecenia_wynik['Suma_kontaktow'] * 3

    return zlecenia_wynik, df_unikalne


@st.cache_data(show_spinner=False)
def wczytaj_gmc121(plik_excel):
    """Zliczanie próbek do SAP z arkusza GM C121 (Preassembly).
    Każdy wiersz z pomiarami = 3 sztuki fizyczne (1 zestaw × 3 szt.).
    """
    df = _wczytaj_arkusz(plik_excel, 'GM C121 (Preassembly)')

    c_zlec    = _znajdz_kolumne(df, 'numer zlecenia')
    c_kontakt = _znajdz_kolumne(df, 'numer kontaktu')
    c_typ     = _znajdz_kolumne(df, 'typ pojedy')
    c_data    = _znajdz_kolumne(df, 'data i czas startu')
    c_zryw    = _znajdz_kolumne(df, 'zaciskania', 'zrywy')
    c_mikro   = _znajdz_kolumne(df, 'zaciskania', 'mikrografie')
    c_poglad  = _znajdz_kolumne(df, 'zaciskania', 'poglądowa')

    df[c_zlec] = df[c_zlec].astype(str).str.strip().replace(['nan', 'None', 'nan.0'], '')
    df['Zlecenie'] = df[c_zlec].replace('', pd.NA).ffill()
    df['Numer kontaktu'] = df[c_kontakt].astype(str).str.strip()
    df['Typ'] = df[c_typ].fillna('').astype(str).str.strip()
    df['Data'] = pd.to_datetime(df[c_data], errors='coerce').ffill()
    df['Data_kalendarzowa'] = df['Data'].dt.date
    df['NR LINII '] = 'GM C121'

    df['_zryw']  = pd.to_numeric(df[c_zryw], errors='coerce')
    df['_mikro'] = pd.to_numeric(df[c_mikro], errors='coerce')
    df['_pogl']  = pd.to_numeric(df[c_poglad], errors='coerce')

    # Liczba kontaktów per wiersz = liczba wypełnionych kolumn wysokości
    # ZMP (zrywy+mikrograf+poglądowa) = 3 kontakty = 3 szt.
    # Z__ / __P (jedna kolumna)       = 1 kontakt  = 1 szt.
    df['n_kontaktow'] = df[['_zryw', '_mikro', '_pogl']].notna().sum(axis=1)
    df['ma_pomiar'] = df['n_kontaktow'] > 0

    df_pom = df[df['ma_pomiar']].copy()

    if df_pom.empty:
        return pd.DataFrame(), pd.DataFrame()

    zlec = df_pom.groupby('Zlecenie').agg(
        Suma_kontaktow=('n_kontaktow', 'sum'),
        Data_max=('Data', 'max'),
    ).reset_index().rename(columns={'Zlecenie': 'Numer zlecenia '})
    zlec['NR LINII '] = 'GM C121'
    zlec['Data i czas startu zlecenia'] = zlec.pop('Data_max')
    zlec['Data_kalendarzowa'] = zlec['Data i czas startu zlecenia'].dt.date
    zlec['Ilosc_zestawow'] = 1
    zlec['Ilosc_fizycznych_probek'] = zlec['Suma_kontaktow']  # 1 kolumna = 1 szt.

    # SAP per PN + Typ
    df_pom['PN'] = df_pom['Numer kontaktu'].astype(str).str.strip()
    df_pom.loc[df_pom['PN'].isin(['nan', 'None', '']), 'PN'] = 'BRAK_PN'
    sap = df_pom.groupby(['PN', 'Typ']).agg(
        Kontakty=('n_kontaktow', 'sum')
    ).reset_index()
    sap['Sztuki'] = sap['Kontakty']
    sap = sap.drop(columns='Kontakty').rename(
        columns={'PN': 'Numer kontaktu (PN)'}
    ).sort_values('Sztuki', ascending=False).reset_index(drop=True)

    return zlec, sap


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
    st.markdown('### 📁 Q-Gate Dashboard')
    plik = st.file_uploader(
        'Wgraj plik Q-Gate:',
        type=['xlsx', 'xlsm'],
        help='Format: xlsx lub xlsm, arkusz "Qgate 2026"',
        label_visibility='collapsed',
    )
    if plik:
        st.success(f'✅ {plik.name}', icon=None)
    else:
        st.caption('Obsługiwane: .xlsx, .xlsm')
    st.caption('wersja 2026.07.23 · elastyczne kolumny')

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
# Inicjalizacja zmiennych filtrów (będą wypełnione w sidebar)
s_dt = e_dt = pd.Timestamp.now()
wybrany_m = ''

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
                    df_f[['Zlecenie', 'NR LINII ', 'Data_kalendarzowa',
                           'Suma_kontaktow', 'Ilosc_fizycznych_probek']]
                    .rename(columns={
                        'NR LINII ': 'Linia',
                        'Data_kalendarzowa': 'Data', 'Suma_kontaktow': 'PN',
                        'Ilosc_fizycznych_probek': 'Szt. SAP',
                    }).reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                )

        # ── GM C121 (Preassembly) ──────────────────────────────────────────────
        st.divider()
        st.markdown('### 🔩 GM C121 (Preassembly) — próbki do SAP')
        try:
            zlec_gm, sap_gm = wczytaj_gmc121(plik)
            if not zlec_gm.empty:
                # Filtruj do tego samego okresu co raport główny
                if tryb == 'Dzienny / Zmianowy':
                    zlec_gm_f = zlec_gm[
                        (zlec_gm['Data i czas startu zlecenia'] >= s_dt) &
                        (zlec_gm['Data i czas startu zlecenia'] <= e_dt)
                    ]
                    sap_gm_f_kontakty = zlec_gm_f  # do filtrowania SAP poniżej
                elif tryb == 'Miesięczny':
                    zlec_gm['YM'] = zlec_gm['Data i czas startu zlecenia'].dt.strftime('%Y-%m')
                    zlec_gm_f = zlec_gm[zlec_gm['YM'] == wybrany_m]
                else:
                    zlec_gm_f = zlec_gm[
                        (zlec_gm['Data i czas startu zlecenia'] >= s_dt) &
                        (zlec_gm['Data i czas startu zlecenia'] < e_dt)
                    ]

                if not zlec_gm_f.empty:
                    cg1, cg2, cg3 = st.columns(3)
                    cg1.metric('🔥 Próbki SAP (GM)', f"{int(zlec_gm_f['Ilosc_fizycznych_probek'].sum()):,} szt.")  # 1 kontakt = 1 szt.
                    cg2.metric('📦 Zestawy (GM)', f"{int(zlec_gm_f['Ilosc_zestawow'].sum()):,} szt.")
                    cg3.metric('🔌 Unikalne PN (GM)', str(int(zlec_gm_f['Suma_kontaktow'].sum())))

                    st.markdown('**Ściągawka SAP — GM C121**')
                    # Filtruj SAP do zleceń z wybranego okresu
                    zlecenia_okresu = set(zlec_gm_f['Numer zlecenia '].tolist())
                    # Przefiltruj df_pom przez zlecenia okresu
                    st.dataframe(
                        sap_gm.rename(columns={
                            'Numer kontaktu (PN)': 'PN',
                        }),
                        use_container_width=True,
                        hide_index=True, height=250,
                    )
                else:
                    st.info('Brak danych GM C121 dla wybranego okresu.')
            else:
                st.info('Brak danych GM C121 w pliku.')
        except Exception as e:
            st.warning(f'Nie udało się wczytać GM C121: {e}')

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
