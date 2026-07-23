"""
Moduł SPC dla Q-Gate Dashboard
Karty Xbar-R — Gen2 (L1-3), Gen3 (L9-10), HPC Uncooled, GM C121 (Preassembly)
"""

import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re

# ─── STAŁE SPC ────────────────────────────────────────────────────────────────

SPC_STALE = {
    2: {'A2': 1.880, 'D3': 0.000, 'D4': 3.267, 'd2': 1.128},
    3: {'A2': 1.023, 'D3': 0.000, 'D4': 2.574, 'd2': 1.693},
    4: {'A2': 0.729, 'D3': 0.000, 'D4': 2.282, 'd2': 2.059},
    5: {'A2': 0.577, 'D3': 0.000, 'D4': 2.114, 'd2': 2.326},
}

# Klucze pomiarowe (używane wewnętrznie, niezależne od nazw w arkuszu)
M_ZRYWY = 'zrywy'
M_MIKRO = 'mikrograf'
M_POGLAD = 'pogladowa'
POMIARY = [M_ZRYWY, M_MIKRO, M_POGLAD]

LINIE_SPC   = ['1', '2', '3', '8', '9', '10', 'HPC UNCOOLED', 'GM C121']
BLOK_NAZWA  = {1: 'Connector', 2: 'Plug'}
NAZWY_LINII = {
    '1': 'Linia 1 (Gen2)', '2': 'Linia 2 (Gen2)', '3': 'Linia 3 (Gen2)',
    '8': 'Linia 8 (Gen2)',
    '9': 'Linia 9 (Gen3)', '10': 'Linia 10 (Gen3)',
    'HPC UNCOOLED': 'HPC 365 Uncooled',
    'GM C121': 'GM C121 (Preassembly)',
}


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
    """Wczytuje arkusz automatycznie wykrywając wiersz nagłówka (2 lub 3).
    W nowszych plikach Qgate 2026 nagłówek przesunął się do wiersza 3."""
    for hdr in (3, 2):
        df = pd.read_excel(plik_excel, sheet_name=sheet_name,
                           header=hdr, engine='openpyxl').copy()
        # Sprawdź czy nagłówek jest sensowny (są prawdziwe nazwy, nie 'Kolumna1')
        cols_str = ' '.join(str(c) for c in df.columns[:12])
        if 'Numer zlecenia' in cols_str and 'Kolumna1' not in cols_str:
            return df
    # Fallback — domyślnie wiersz 2
    return pd.read_excel(plik_excel, sheet_name=sheet_name,
                         header=2, engine='openpyxl').copy()

# ─── PALETA CIEMNA ────────────────────────────────────────────────────────────

BG_PLOT  = '#0f172a'
BG_PAPER = '#1e293b'
COL_GRID = '#1e3a5f'
COL_TICK = '#94a3b8'
COL_TITLE= '#e2e8f0'
COL_XBAR = '#38bdf8'
COL_R    = '#34d399'
COL_UCL  = '#f87171'
COL_CL   = '#38bdf8'
COL_USL  = '#fbbf24'
COL_OUT  = '#f87171'
COL_WARN = '#fb923c'

ZONE_COLORS = [
    (3, 'rgba(239,68,68,0.07)'),
    (2, 'rgba(251,146,60,0.09)'),
    (1, 'rgba(34,197,94,0.12)'),
]
WEEK_BAND_A = 'rgba(56,189,248,0.04)'
WEEK_BAND_B = 'rgba(15,23,42,0.0)'

# ─── NORMALIZACJA ─────────────────────────────────────────────────────────────

def _norm_linia(t):
    if pd.isna(t): return ''
    t = str(t).upper().strip()
    t = re.sub(r'[\.\-/_]', ' ', t)
    t = ' '.join(t.split())
    # NACS DC osobno — sprawdzamy PRZED ogólnym NACS
    if 'NACS' in t and 'DC' in t: return 'NACS DC'
    if t == 'NAC' or t == 'NACS': return 'NACS'
    # Uni inlet: wszystkie warianty (z/bez spacji) → INLET
    if 'UNIINLET' in t.replace(' ', '') or t == 'INLET': return 'INLET'
    if 'CCSD' in t.replace(' ', '') or t == 'CCS D': return 'CCS D'
    if 'HPC' in t and 'UNCOOL' in t: return 'HPC UNCOOLED'
    if 'HPC' in t and '1' in t: return 'HPC 1.0'
    if 'HPC' in t and '2' in t: return 'HPC 2.0'
    if 'GM' in t and 'C121' in t: return 'GM C121'
    if 'PREASSY' in t and 'GEN 3' in t: return 'PREASSY GEN 3'
    if 'PREASSY' in t and ('GEN 2' in t or 'GEN2' in t): return 'PREASSY GEN 2'
    return t

# ─── WCZYTANIE DANYCH — QGATE 2026 ───────────────────────────────────────────

def _wczytaj_qgate(plik_excel) -> pd.DataFrame:
    df = _wczytaj_arkusz(plik_excel, 'Qgate 2026')

    # Wykryj kolumny elastycznie (odporność na spacje/wersje pliku)
    c_zlec    = _znajdz_kolumne(df, 'numer zlecenia')
    c_linia   = _znajdz_kolumne(df, 'nr linii')
    c_kontakt = _znajdz_kolumne(df, 'numer kontaktu')
    c_typ     = _znajdz_kolumne(df, 'typ pojedy')
    c_masz    = _znajdz_kolumne(df, 'nazwa maszyny')
    c_przek   = _znajdz_kolumne(df, 'przekrój')
    c_data    = _znajdz_kolumne(df, 'data i czas startu')
    c_lsl     = _znajdz_kolumne(df, 'lsl')
    c_usl     = _znajdz_kolumne(df, 'usl')
    c_zryw    = _znajdz_kolumne(df, 'zaciskania', 'zrywy')
    c_mikro   = _znajdz_kolumne(df, 'zaciskania', 'mikrografie')
    c_poglad  = _znajdz_kolumne(df, 'zaciskania', 'poglądowa')

    for col in [c_zlec, c_linia, c_kontakt, c_typ]:
        df[col] = df[col].astype(str).str.strip().replace(
            ['nan', 'None', 'nan.0'], '')

    df['Zlecenie'] = df[c_zlec].replace('', pd.NA).ffill()
    df['LiniaRaw'] = df[c_linia].replace('', pd.NA).ffill()
    df['Linia']    = df['LiniaRaw'].apply(_norm_linia)
    df['Typ']      = df[c_typ].astype(str).str.strip()
    df['Maszyna']  = df[c_masz].astype(str).str.strip()
    df['Kontakt']  = df[c_kontakt].astype(str).str.strip()
    df['Przekroj'] = df[c_przek].astype(str).str.strip()
    df['Data']     = pd.to_datetime(df[c_data], errors='coerce').ffill()
    df['LSL']      = pd.to_numeric(df[c_lsl], errors='coerce')
    df['USL']      = pd.to_numeric(df[c_usl], errors='coerce')

    # Ujednolicenie kolumn pomiarowych do kluczy wewnętrznych
    df[M_ZRYWY]  = pd.to_numeric(df[c_zryw], errors='coerce')
    df[M_MIKRO]  = pd.to_numeric(df[c_mikro], errors='coerce')
    df[M_POGLAD] = pd.to_numeric(df[c_poglad], errors='coerce')
    df['ma_pomiar'] = df[POMIARY].notna().any(axis=1)

    df = df[df['Linia'].isin(LINIE_SPC)].copy()

    # Bloki connector/plug przez unikalny kontakt CP
    wyniki = []
    for (z, linia), grp in df.groupby(['Zlecenie', 'Linia'], sort=False):
        grp = grp.reset_index(drop=True)
        cp_rows     = grp[grp['Typ'] == 'CP']
        unikalne_cp = list(dict.fromkeys(cp_rows['Kontakt'].tolist()))
        if not unikalne_cp:
            grp['blok'] = 1
        else:
            cp_to_blok  = {cp: i + 1 for i, cp in enumerate(unikalne_cp)}
            cur = None
            bloki = []
            for _, row in grp.iterrows():
                if row['Typ'] == 'CP':
                    cur = cp_to_blok.get(row['Kontakt'])
                bloki.append(cur)
            grp['blok'] = bloki
        wyniki.append(grp)

    df_b   = pd.concat(wyniki, ignore_index=True) if wyniki else pd.DataFrame()
    df_pom = df_b[df_b['ma_pomiar']].copy()

    podgrupy = []
    for keys, grp in df_pom.groupby(
        ['Zlecenie', 'Linia', 'blok', 'Typ', 'Maszyna', 'Przekroj'],
        sort=False
    ):
        z, linia, blok, typ, masz, przekroj = keys
        vals = grp[POMIARY].values.flatten()
        vals = vals[~np.isnan(vals)]
        if len(vals) == 0:
            continue
        lsl  = grp['LSL'].dropna().iloc[0] if grp['LSL'].notna().any() else np.nan
        usl  = grp['USL'].dropna().iloc[0] if grp['USL'].notna().any() else np.nan
        data = grp['Data'].iloc[0]
        podgrupy.append({
            'Zlecenie': z,     'Linia': linia, 'Blok': blok,
            'Typ': typ,        'Maszyna': masz, 'Przekroj': przekroj,
            'Data': data,      'Tydzien': int(data.isocalendar().week),
            'Xbar': float(np.mean(vals)),
            'R':    float(np.max(vals) - np.min(vals)),
            'n':    len(vals),
            'LSL':  float(lsl) if not np.isnan(lsl) else np.nan,
            'USL':  float(usl) if not np.isnan(usl) else np.nan,
        })
    return pd.DataFrame(podgrupy)


# ─── WCZYTANIE DANYCH — GM C121 ───────────────────────────────────────────────

def _wczytaj_gmc121(plik_excel) -> pd.DataFrame:
    """
    GM C121 (Preassembly) — każdy wiersz z jakimkolwiek wpisem wysokości
    zaciskania (zrywy / mikrograf / poglądowa) = jeden kontakt = podgrupa SPC.
    Logika zliczania SAP jest identyczna dla wszystkich typów:
      - MOLEX/MCON regularne: zrywy+mikrograf+poglądowa w jednym wierszu
      - MOLEX/MCON tygodniowe poglądowe: tylko poglądowa (wszystkie kolory)
      - DC/AC/PE: 5 wierszy per seria (3×zrywy + 1×poglądowa + 1×przekrój)
    Każdy wiersz z wpisem = 1 kontakt do SAP = 3 sztuki fizyczne.
    Wiersze bez żadnego wpisu wysokości = brak kontaktu = 0 sztuk.
    """
    df = _wczytaj_arkusz(plik_excel, 'GM C121 (Preassembly)')

    c_zlec    = _znajdz_kolumne(df, 'numer zlecenia')
    c_kontakt = _znajdz_kolumne(df, 'numer kontaktu')
    c_typ     = _znajdz_kolumne(df, 'typ pojedy')
    c_przek   = _znajdz_kolumne(df, 'przekrój')
    c_data    = _znajdz_kolumne(df, 'data i czas startu')
    c_lsl     = _znajdz_kolumne(df, 'dolna granica tolerancji')
    c_usl     = _znajdz_kolumne(df, 'górna gr', 'tolerancji')
    c_zryw    = _znajdz_kolumne(df, 'zaciskania', 'zrywy')
    c_mikro   = _znajdz_kolumne(df, 'zaciskania', 'mikrografie')
    c_poglad  = _znajdz_kolumne(df, 'zaciskania', 'poglądowa')

    df[c_zlec] = df[c_zlec].astype(str).str.strip().replace(
        ['nan', 'None', 'nan.0'], '')
    df['Zlecenie'] = df[c_zlec].replace('', pd.NA).ffill()
    df['Linia']    = 'GM C121'
    df['Blok']     = 1
    df['Typ']      = df[c_typ].fillna('').astype(str).str.strip()
    df['Maszyna']  = 'GM C121'
    df['Kontakt']  = df[c_kontakt].astype(str).str.strip()
    df['Przekroj'] = df[c_przek].astype(str).str.strip()
    df['Data']     = pd.to_datetime(df[c_data], errors='coerce').ffill()
    df['LSL']      = pd.to_numeric(df[c_lsl], errors='coerce')
    df['USL']      = pd.to_numeric(df[c_usl], errors='coerce')

    df[M_ZRYWY]  = pd.to_numeric(df[c_zryw], errors='coerce')
    df[M_MIKRO]  = pd.to_numeric(df[c_mikro], errors='coerce')
    df[M_POGLAD] = pd.to_numeric(df[c_poglad], errors='coerce')

    # Dla SPC bierzemy TYLKO wiersze z kompletem 3 pomiarów (ZMP, n=3)
    # Wiersze z 1 lub 2 kolumnami (tygodniowe poglądowe, same zrywy) mają n<3
    # i zaburzają granice kontrolne Xbar-R — zostawiamy je dla zliczania SAP,
    # ale nie dla kart kontrolnych.
    df['n_pomiarow'] = df[POMIARY].notna().sum(axis=1)
    df['ma_pomiar'] = df['n_pomiarow'] == 3   # tylko kompletne podgrupy n=3
    df_pom = df[df['ma_pomiar']].copy()

    # Dla SPC: każdy wiersz z wpisami to potencjalna podgrupa.
    # n = liczba wypełnionych kolumn (1, 2 lub 3) — odpowiada liczbie kontaktów.
    # Podgrupy z n=1 (tylko jeden pomiar) też są wartościowe dla kart I-MR,
    # ale dla Xbar-R minimalne n=2. Zostawiamy wszystkie — oblicz_granice
    # samo dobierze właściwe stałe SPC.
    podgrupy = []
    for _, row in df_pom.iterrows():
        data = row['Data']
        if pd.isna(data):
            continue
        vals = np.array([row[p] for p in POMIARY], dtype=float)
        vals_ok = vals[~np.isnan(vals)]
        if len(vals_ok) == 0:
            continue
        typ_val = row['Typ'] if row['Typ'] not in ('', 'nan', 'None') else 'BRAK'
        podgrupy.append({
            'Zlecenie': str(row['Zlecenie']),
            'Linia':    'GM C121',
            'Blok':     1,
            'Typ':      typ_val,
            'Maszyna':  'GM C121',
            'Przekroj': row['Przekroj'],
            'Data':     data,
            'Tydzien':  int(data.isocalendar().week),
            'Xbar':     float(np.mean(vals_ok)),
            'R':        float(np.max(vals_ok) - np.min(vals_ok)) if len(vals_ok) > 1 else 0.0,
            'n':        len(vals_ok),   # 1=tylko zrywy/mikr/poglądowa, 3=komplet
            'LSL':      float(row['LSL']) if not pd.isna(row['LSL']) else np.nan,
            'USL':      float(row['USL']) if not pd.isna(row['USL']) else np.nan,
        })
    return pd.DataFrame(podgrupy)


# ─── GŁÓWNA FUNKCJA WCZYTUJĄCA ────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def przygotuj_dane_spc(plik_excel) -> pd.DataFrame:
    df_q  = _wczytaj_qgate(plik_excel)
    df_gm = _wczytaj_gmc121(plik_excel)
    return pd.concat([df_q, df_gm], ignore_index=True)


# ─── OBLICZENIA ───────────────────────────────────────────────────────────────

def oblicz_granice(df_pg: pd.DataFrame) -> dict:
    if df_pg.empty:
        return {}
    n_t  = int(df_pg['n'].mode().iloc[0])
    n_t  = max(2, min(n_t, 5))
    stk  = SPC_STALE[n_t]

    Xbar = df_pg['Xbar'].mean()
    Rbar = df_pg['R'].mean()

    UCL_x = Xbar + stk['A2'] * Rbar
    LCL_x = Xbar - stk['A2'] * Rbar
    UCL_r = stk['D4'] * Rbar
    LCL_r = stk['D3'] * Rbar

    lsl = df_pg['LSL'].dropna().mean()
    usl = df_pg['USL'].dropna().mean()
    sigma = Rbar / stk['d2'] if stk['d2'] > 0 else np.nan

    cp = cpk = np.nan
    try:
        if not any(np.isnan(v) for v in [lsl, usl, sigma]) and sigma > 0:
            cp  = (usl - lsl) / (6 * sigma)
            cpk = min((usl - Xbar) / (3 * sigma), (Xbar - lsl) / (3 * sigma))
    except Exception:
        pass

    return {
        'Xbarbar': Xbar, 'Rbar': Rbar,
        'UCL_x': UCL_x, 'LCL_x': LCL_x,
        'UCL_r': UCL_r, 'LCL_r': LCL_r,
        'LSL': lsl, 'USL': usl,
        'sigma_hat': sigma,
        'Cp': cp, 'Cpk': cpk,
        'n_podgrup': len(df_pg),
    }


def wykryj_sygnaly(df_pg: pd.DataFrame, granice: dict) -> pd.Series:
    if df_pg.empty or not granice:
        return pd.Series([], dtype=str)
    x  = df_pg['Xbar'].values
    cl = granice['Xbarbar']
    s3 = granice['UCL_x'] - cl
    s1 = s3 / 3
    s2 = 2 * s1
    out = np.full(len(x), '', dtype=object)

    for i in range(len(x)):
        if abs(x[i] - cl) > s3:
            out[i] = 'poza_3sigma'; continue
        if i >= 2:
            seg = x[i-2:i+1]
            if (sum(v > cl + s2 for v in seg) >= 2 or
                    sum(v < cl - s2 for v in seg) >= 2):
                out[i] = 'reguła_2z3'; continue
        if i >= 4:
            seg = x[i-4:i+1]
            if (sum(v > cl + s1 for v in seg) >= 4 or
                    sum(v < cl - s1 for v in seg) >= 4):
                out[i] = 'reguła_4z5'; continue
        if i >= 7:
            seg = x[i-7:i+1]
            if all(v > cl for v in seg) or all(v < cl for v in seg):
                out[i] = 'reguła_8'

    return pd.Series(out, index=df_pg.index)


# ─── RYSOWANIE KARTY ─────────────────────────────────────────────────────────

def rysuj_karte_xbar_r(df_pg: pd.DataFrame, granice: dict, tytul: str) -> go.Figure:
    if df_pg.empty:
        return go.Figure()

    df_pg = df_pg.sort_values('Data').reset_index(drop=True)
    df_pg['sygnal'] = wykryj_sygnaly(df_pg, granice).values

    x_idx  = list(range(len(df_pg)))
    cl     = granice['Xbarbar']
    ucl_x  = granice['UCL_x']
    lcl_x  = granice['LCL_x']
    rbar   = granice['Rbar']
    ucl_r  = granice['UCL_r']
    sigma3 = ucl_x - cl
    sigma1 = sigma3 / 3
    lsl    = granice.get('LSL', np.nan)
    usl    = granice.get('USL', np.nan)

    labels = [
        f"{r['Data'].strftime('%d.%m')} #{str(r['Zlecenie'])[-4:]}"
        for _, r in df_pg.iterrows()
    ]

    # subplot_titles jako nagłówki — nie kolidują z hlines
    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=(
            'Karta X̄ — Średnie podgrup (3 pomiary wysokości zaciskania per zlecenie)',
            'Karta R — Rozstępy podgrup  (max − min;  mały R = powtarzalny proces)',
        ),
        vertical_spacing=0.14,
        row_heights=[0.62, 0.38],
        shared_xaxes=True,
    )

    # Styl nagłówków podwykresów
    for ann in fig.layout.annotations:
        ann.update(font=dict(size=12, color='#94a3b8'), x=0, xanchor='left')

    # ── Pasy tygodni ──────────────────────────────────────────────────────────
    week_bounds, prev_w, seg_start = [], None, 0
    for i, (_, row) in enumerate(df_pg.iterrows()):
        w = row['Tydzien']
        if prev_w is None:
            seg_start = i
        elif w != prev_w:
            week_bounds.append((seg_start, i - 1, prev_w))
            seg_start = i
        prev_w = w
    if prev_w is not None:
        week_bounds.append((seg_start, len(df_pg) - 1, prev_w))

    for bi, (i0, i1, tydzien) in enumerate(week_bounds):
        color = WEEK_BAND_A if bi % 2 == 0 else WEEK_BAND_B
        for rn in [1, 2]:
            fig.add_vrect(x0=i0 - 0.5, x1=i1 + 0.5, fillcolor=color,
                          line_width=0, layer='below', row=rn, col=1)
        # Etykieta tygodnia nad górnym wykresem tylko
        fig.add_annotation(
            x=(i0 + i1) / 2, y=1.04, yref='paper',
            text=f'<b>Tydz.{tydzien}</b>',
            showarrow=False, xanchor='center',
            font=dict(size=8, color='#475569'),
        )

    for i0, _, _ in week_bounds[1:]:
        for rn in [1, 2]:
            fig.add_vline(x=i0 - 0.5,
                          line_color='rgba(100,116,139,0.3)',
                          line_dash='dot', line_width=1,
                          row=rn, col=1)

    # ── Strefy sigma ──────────────────────────────────────────────────────────
    for mult, color in ZONE_COLORS:
        y_top = [cl + mult * sigma1] * len(x_idx)
        y_bot = [cl - mult * sigma1] * len(x_idx)
        fig.add_trace(go.Scatter(
            x=x_idx + x_idx[::-1], y=y_top + y_bot[::-1],
            fill='toself', fillcolor=color,
            line=dict(width=0), showlegend=False,
            hoverinfo='skip', mode='lines',
        ), row=1, col=1)

    # ── USL / LSL — po lewej żeby nie kolidowały z UCL/LCL po prawej ─────────
    for val, name in [(usl, 'USL'), (lsl, 'LSL')]:
        try:
            fv = float(val)
            if not np.isnan(fv):
                fig.add_hline(
                    y=fv, line_color=COL_USL,
                    line_dash='dot', line_width=2,
                    annotation_text=f'<b>{name}</b>={fv:.3f}',
                    annotation_position='left',         # ← lewa strona
                    annotation_font=dict(color=COL_USL, size=11),
                    row=1, col=1,
                )
        except (TypeError, ValueError):
            pass

    # ── UCL / X̄̄ / LCL — po prawej ────────────────────────────────────────────
    for val, name, col, dash, width in [
        (ucl_x, 'UCL', COL_UCL, 'dash',  2.0),
        (lcl_x, 'LCL', COL_UCL, 'dash',  2.0),
        (cl,    'X̄̄',  COL_CL,  'solid', 2.5),
    ]:
        fig.add_hline(
            y=val, line_color=col, line_dash=dash, line_width=width,
            annotation_text=f'<b>{name}</b>={val:.4f}',
            annotation_position='right',
            annotation_font=dict(color=col, size=11),
            row=1, col=1,
        )

    # ── Hover ─────────────────────────────────────────────────────────────────
    hover_x = [
        (f"<b>Zlecenie:</b> {r['Zlecenie']}<br>"
         f"<b>Data:</b> {r['Data'].strftime('%d.%m.%Y %H:%M')}<br>"
         f"<b>X̄</b> = {r['Xbar']:.4f} mm<br>"
         f"<b>R</b> = {r['R']:.4f} mm   n={r['n']}   Tydz.{r['Tydzien']}")
        for _, r in df_pg.iterrows()
    ]

    # ── Linia łącząca punkty ──────────────────────────────────────────────────
    fig.add_trace(go.Scatter(
        x=x_idx, y=df_pg['Xbar'], mode='lines',
        line=dict(color='rgba(148,163,184,0.3)', width=1.2),
        showlegend=False, hoverinfo='skip',
    ), row=1, col=1)

    # ── Punkty X̄ ──────────────────────────────────────────────────────────────
    mask_ok   = df_pg['sygnal'] == ''
    mask_warn = df_pg['sygnal'].isin(['reguła_2z3', 'reguła_4z5', 'reguła_8'])
    mask_out  = df_pg['sygnal'] == 'poza_3sigma'

    for mask, color, size, symbol, lname in [
        (mask_ok,   COL_XBAR, 8,  'circle',     'W kontroli'),
        (mask_warn, COL_WARN, 11, 'diamond',     'Ostrzeżenie (reguła Nelson)'),
        (mask_out,  COL_OUT,  13, 'x-thin-open', 'Poza UCL/LCL'),
    ]:
        if not mask.any():
            continue
        idxs = [i for i, m in enumerate(mask) if m]
        fig.add_trace(go.Scatter(
            x=idxs, y=df_pg.loc[mask, 'Xbar'],
            mode='markers',
            marker=dict(color=color, size=size, symbol=symbol,
                        line=dict(color='white', width=1.2)),
            name=lname,
            customdata=[hover_x[i] for i in idxs],
            hovertemplate='%{customdata}<extra></extra>',
        ), row=1, col=1)

    # ── Karta R ───────────────────────────────────────────────────────────────
    # Strefa normalna pod UCL_r
    fig.add_trace(go.Scatter(
        x=x_idx + x_idx[::-1],
        y=[ucl_r] * len(x_idx) + [0] * len(x_idx),
        fill='toself', fillcolor='rgba(52,211,153,0.06)',
        line=dict(width=0), showlegend=False,
        hoverinfo='skip', mode='lines',
    ), row=2, col=1)

    fig.add_hline(y=ucl_r, line_color=COL_UCL, line_dash='dash', line_width=1.8,
                  annotation_text=f'<b>UCL</b>={ucl_r:.4f}',
                  annotation_position='right',
                  annotation_font=dict(color=COL_UCL, size=11),
                  row=2, col=1)
    fig.add_hline(y=rbar, line_color=COL_R, line_dash='solid', line_width=2.0,
                  annotation_text=f'<b>R̄</b>={rbar:.4f}',
                  annotation_position='right',
                  annotation_font=dict(color=COL_R, size=11),
                  row=2, col=1)

    mask_r_out = df_pg['R'] > ucl_r
    hover_r = [
        (f"<b>R</b> = {r['R']:.4f} mm<br>"
         f"<b>Zlecenie:</b> {r['Zlecenie']}<br>"
         f"<b>Data:</b> {r['Data'].strftime('%d.%m.%Y')}")
        for _, r in df_pg.iterrows()
    ]
    fig.add_trace(go.Bar(
        x=x_idx, y=df_pg['R'],
        name='Rozstęp R',
        marker=dict(
            color=['#f87171' if m else COL_R for m in mask_r_out],
            opacity=0.78, line=dict(width=0),
        ),
        customdata=hover_r,
        hovertemplate='%{customdata}<extra></extra>',
    ), row=2, col=1)

    # ── Osie ──────────────────────────────────────────────────────────────────
    n_pts = len(df_pg)
    step  = max(1, n_pts // 20)
    tickvals = list(range(0, n_pts, step))

    fig.update_xaxes(showticklabels=False, showgrid=False, zeroline=False, row=1, col=1)
    fig.update_xaxes(
        tickvals=tickvals, ticktext=[labels[i] for i in tickvals],
        tickfont=dict(size=8, color=COL_TICK), tickangle=-35,
        showgrid=False, zeroline=False, row=2, col=1,
    )
    fig.update_yaxes(
        gridcolor=COL_GRID, gridwidth=1,
        tickfont=dict(size=9, color=COL_TICK), zeroline=False,
        title_text='mm', title_font=dict(color=COL_TICK, size=10),
    )

    # ── Layout ────────────────────────────────────────────────────────────────
    fig.update_layout(
        title=dict(
            text=f'<b>{tytul}</b>',
            font=dict(size=13, color=COL_TITLE),
            x=0, xanchor='left', pad=dict(t=4),
        ),
        height=670,
        margin=dict(l=60, r=120, t=100, b=60),
        plot_bgcolor=BG_PLOT, paper_bgcolor=BG_PAPER,
        font=dict(color=COL_TICK, family='Inter, Arial, sans-serif'),
        legend=dict(
            orientation='h', y=-0.10, x=0.5, xanchor='center',
            font=dict(size=11, color=COL_TITLE),
            bgcolor='rgba(15,23,42,0.7)',
            bordercolor=COL_GRID, borderwidth=1,
            traceorder='normal',
        ),
        hovermode='x unified',
        hoverlabel=dict(
            bgcolor='#0f172a', bordercolor='#334155',
            font=dict(color='#e2e8f0', size=12), namelength=-1,
        ),
        bargap=0.18,
        # Usuń duplikaty z legendy
        showlegend=True,
    )
    return fig


# ─── METRYKI Cp/Cpk ───────────────────────────────────────────────────────────

def pokaz_metryki_spc(granice: dict):
    cp    = granice.get('Cp',  np.nan)
    cpk   = granice.get('Cpk', np.nan)
    n     = granice.get('n_podgrup', 0)
    sigma = granice.get('sigma_hat', np.nan)

    def kolor(v):
        if np.isnan(v): return '#6b7280'
        if v >= 1.67:   return '#34d399'
        if v >= 1.33:   return '#38bdf8'
        if v >= 1.00:   return '#fbbf24'
        return '#f87171'

    def status(v):
        if np.isnan(v): return '—'
        if v >= 1.67:   return '✅ Doskonały  (≥1.67)'
        if v >= 1.33:   return '✅ Zdolny     (≥1.33)'
        if v >= 1.00:   return '⚠️ Marginalny (≥1.00)'
        return           '❌ Niezdolny  (<1.00)'

    C = ("background:#0f172a;border-left:4px solid {bc};"
         "padding:13px 16px;border-radius:10px;"
         "box-shadow:0 0 14px {glow};")
    L = "font-size:0.72rem;color:#64748b;text-transform:uppercase;letter-spacing:.07em;margin-bottom:5px;"
    V = "font-size:2.1rem;font-weight:700;line-height:1;color:{vc};"
    S = "font-size:0.73rem;color:#475569;margin-top:4px;"

    col1, col2, col3 = st.columns(3)

    with col1:
        bc = '#38bdf8'; glow = 'rgba(56,189,248,.14)'
        st.markdown(f"""<div style="{C.format(bc=bc,glow=glow)}">
            <div style="{L}">Cp — zdolność potencjalna</div>
            <div style="{V.format(vc=bc)}">{f'{cp:.3f}' if not np.isnan(cp) else '—'}</div>
            <div style="{S}">symetryczna względem środka tolerancji</div>
        </div>""", unsafe_allow_html=True)

    with col2:
        vc = kolor(cpk)
        glow = 'rgba(248,113,113,.14)' if vc == '#f87171' else 'rgba(52,211,153,.12)'
        st.markdown(f"""<div style="{C.format(bc=vc,glow=glow)}">
            <div style="{L}">Cpk — zdolność rzeczywista</div>
            <div style="{V.format(vc=vc)}">{f'{cpk:.3f}' if not np.isnan(cpk) else '—'}</div>
            <div style="{S}">{status(cpk)}</div>
        </div>""", unsafe_allow_html=True)

    with col3:
        bc = '#818cf8'; glow = 'rgba(129,140,248,.14)'
        xbar = f"{granice.get('Xbarbar', 0):.4f}"
        sig  = f"{sigma:.4f}" if not np.isnan(sigma) else '—'
        st.markdown(f"""<div style="{C.format(bc=bc,glow=glow)}">
            <div style="{L}">Podgrupy / σ̂ procesu</div>
            <div style="{V.format(vc=bc)}">{n}</div>
            <div style="{S}">X̄̄ = {xbar} mm &nbsp;·&nbsp; σ̂ = {sig} mm</div>
        </div>""", unsafe_allow_html=True)


# ─── GŁÓWNY WIDOK ─────────────────────────────────────────────────────────────

def pokaz_spc(plik_excel):
    st.subheader('📈 Karty kontrolne SPC — Xbar-R')

    with st.spinner('Przygotowuję dane SPC...'):
        df_pg = przygotuj_dane_spc(plik_excel)

    if df_pg.empty:
        st.warning('Brak danych pomiarowych.')
        return

    # ── Filtry — wiersz 1 ─────────────────────────────────────────────────────
    col_f1, col_f2, col_f3 = st.columns(3)

    with col_f1:
        dostepne = sorted(df_pg['Linia'].unique(),
                          key=lambda x: list(LINIE_SPC).index(x) if x in LINIE_SPC else 99)
        wybrana_linia = st.selectbox('Linia', dostepne,
                                     format_func=lambda x: NAZWY_LINII.get(x, x))

    df_L = df_pg[df_pg['Linia'] == wybrana_linia]

    with col_f2:
        bloki = sorted([b for b in df_L['Blok'].dropna().unique() if b is not None], key=lambda x: float(x))
        if len(bloki) > 1:
            wybrany_blok = st.selectbox(
                'Strona kabla', bloki,
                format_func=lambda x: BLOK_NAZWA.get(int(x), f'Blok {x}') if x else '—')
        else:
            wybrany_blok = bloki[0] if bloki else None
            label = BLOK_NAZWA.get(int(wybrany_blok), '—') if wybrany_blok else '—'
            st.selectbox('Strona kabla', [label], disabled=True)

    df_B = df_L[df_L['Blok'] == wybrany_blok] if wybrany_blok else df_L

    with col_f3:
        typy = sorted([t for t in df_B['Typ'].dropna().unique() if str(t) not in ('nan','None','')])
        wybrany_typ = st.selectbox('Typ przewodu / kontaktu', typy)

    df_T = df_B[df_B['Typ'] == wybrany_typ]

    # ── Filtry — wiersz 2 ─────────────────────────────────────────────────────
    col_p, col_t = st.columns([1, 2])

    with col_p:
        przekroje = sorted([p for p in df_T['Przekroj'].dropna().unique() if str(p) not in ('nan','None','')])
        if len(przekroje) > 1:
            wybrany_przekroj = st.selectbox('Przekrój (mm²)', przekroje)
        else:
            wybrany_przekroj = przekroje[0] if przekroje else None
            st.selectbox('Przekrój (mm²)', [wybrany_przekroj or '—'], disabled=True)

    df_final = df_T[df_T['Przekroj'] == wybrany_przekroj].copy() \
               if wybrany_przekroj else df_T.copy()

    with col_t:
        # Sortuj po dacie żeby tygodnie z przełomu roku były w dobrej kolejności
        tydz_daty = (df_final.groupby('Tydzien')['Data'].min()
                     .reset_index().sort_values('Data'))
        tygodnie = [int(t) for t in tydz_daty['Tydzien'].tolist()]
        if len(tygodnie) > 1:
            zakres = st.select_slider(
                'Zakres tygodni (ISO)',
                options=tygodnie,
                value=(tygodnie[0], tygodnie[-1]),
            )
            # Filtruj po dacie a nie numerze tygodnia
            min_data = tydz_daty[tydz_daty['Tydzien']==zakres[0]]['Data'].iloc[0]
            max_data = (df_final[df_final['Tydzien']==zakres[1]]['Data'].max()
                        + pd.Timedelta(days=7))
            df_final = df_final[
                (df_final['Data'] >= min_data) & (df_final['Data'] <= max_data)
            ]

    if df_final.empty:
        st.info('Brak danych dla wybranej kombinacji filtrów.')
        return

    # ── Obliczenia i wykresy ───────────────────────────────────────────────────
    granice = oblicz_granice(df_final)

    blok_str    = BLOK_NAZWA.get(int(wybrany_blok), '') if wybrany_blok else ''
    linia_str   = NAZWY_LINII.get(wybrana_linia, wybrana_linia)
    przekroj_str = wybrany_przekroj or ''
    # GM C121 nie ma sensu pokazywać bloku ani maszyny w tytule
    if wybrana_linia == 'GM C121':
        tytul = f'{linia_str} | {wybrany_typ}  {przekroj_str}'
    else:
        tytul = f'{linia_str} | {blok_str} | {wybrany_typ}  {przekroj_str}'

    pokaz_metryki_spc(granice)
    st.markdown('<br>', unsafe_allow_html=True)

    fig = rysuj_karte_xbar_r(df_final, granice, tytul)
    st.plotly_chart(fig, use_container_width=True)

    # Tabela danych
    with st.expander('📋 Dane podgrup', expanded=False):
        df_show = df_final[['Data', 'Zlecenie', 'Tydzien', 'Xbar', 'R', 'n', 'LSL', 'USL']].copy()
        df_show['Data']  = df_show['Data'].dt.strftime('%d.%m.%Y %H:%M')
        df_show['Xbar']  = df_show['Xbar'].round(4)
        df_show['R']     = df_show['R'].round(4)
        df_show = df_show.rename(columns={
            'Tydzien': 'Tydz.', 'Xbar': 'X̄ (mm)',
            'R': 'R (mm)', 'n': 'n', 'Zlecenie': 'Nr zlecenia',
        })
        st.dataframe(df_show.reset_index(drop=True),
                     use_container_width=True, hide_index=True)

    # Legenda reguł — raz, na dole
    with st.expander('ℹ️ Reguły sygnałowe Nelson', expanded=False):
        st.markdown("""
| Symbol | Reguła | Co oznacza |
|--------|--------|------------|
| 🔵 ● | W kontroli | Punkt w granicach ±3σ, brak wzorców |
| 🟠 ◆ | 2 z 3 poza 2σ | 2 z 3 kolejnych po tej samej stronie za 2σ |
| 🟠 ◆ | 4 z 5 poza 1σ | 4 z 5 kolejnych po tej samej stronie za 1σ |
| 🟠 ◆ | 8 po jednej stronie | 8 kolejnych po tej samej stronie X̄̄ |
| 🔴 ✕ | Poza 3σ | Punkt poza UCL lub LCL — działaj natychmiast |
| 🟡 ··· | USL / LSL | Granice specyfikacji z arkusza (po lewej stronie wykresu) |
""")
