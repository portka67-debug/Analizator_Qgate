"""
Moduł SPC dla Q-Gate Dashboard
Karty kontrolne Xbar-R dla linii Gen2 (1,2,3), Gen3 (9,10), HPC Uncooled
"""

import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re

# ─── STAŁE SPC ───────────────────────────────────────────────────────────────

# Stałe dla karty Xbar-R (n=3)
# A2, D3, D4 wg normy ISO/AIAG
SPC_STALE = {
    2: {'A2': 1.880, 'D3': 0.000, 'D4': 3.267, 'd2': 1.128},
    3: {'A2': 1.023, 'D3': 0.000, 'D4': 2.574, 'd2': 1.693},
    4: {'A2': 0.729, 'D3': 0.000, 'D4': 2.282, 'd2': 2.059},
    5: {'A2': 0.577, 'D3': 0.000, 'D4': 2.114, 'd2': 2.326},
}

POMIARY = [
    'Wysokość zaciskania - zrywy',
    'Wysokość zaciskania -  kontrola przekroju na mikrografie',
    'Wysokość zaciskania - próbka poglądowa',
]

LINIE_SPC = ['1', '2', '3', '9', '10', 'HPC UNCOOLED']

NAZWY_LINII = {
    '1': 'Linia 1 (Gen2)',
    '2': 'Linia 2 (Gen2)',
    '3': 'Linia 3 (Gen2)',
    '9': 'Linia 9 (Gen3)',
    '10': 'Linia 10 (Gen3)',
    'HPC UNCOOLED': 'HPC 365 Uncooled',
}

BLOK_NAZWA = {1: 'Connector', 2: 'Plug'}

# Kolory
COL_XBAR = '#2563eb'
COL_R    = '#059669'
COL_UCL  = '#dc2626'
COL_LCL  = '#dc2626'
COL_CL   = '#6b7280'
COL_USL  = '#f59e0b'
COL_LSL  = '#f59e0b'
COL_OUT  = '#ef4444'
COL_WARN = '#f97316'


# ─── NORMALIZACJA LINII ───────────────────────────────────────────────────────

def norm_linia(t):
    if pd.isna(t):
        return ''
    t = str(t).upper().strip()
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


# ─── WCZYTANIE I PRZYGOTOWANIE DANYCH ────────────────────────────────────────

@st.cache_data(show_spinner=False)
def przygotuj_dane_spc(plik_excel):
    df = pd.read_excel(plik_excel, sheet_name='Qgate 2026', header=2, engine='openpyxl').copy()

    # Czyszczenie podstawowe
    for col in ['Numer zlecenia ', 'NR LINII ', 'Numer kontaktu', 'Typ pojedyńczego przewodu']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace(['nan', 'None', 'nan.0'], '')

    df['Numer zlecenia '] = df['Numer zlecenia '].replace('', pd.NA).ffill()
    df['NR LINII '] = df['NR LINII '].replace('', pd.NA).ffill()
    df['Linia'] = df['NR LINII '].apply(norm_linia)
    df['Typ'] = df['Typ pojedyńczego przewodu'].astype(str).str.strip()
    df['Maszyna'] = df['Nazwa maszyny'].astype(str).str.strip()
    df['Numer kontaktu'] = df['Numer kontaktu'].astype(str).str.strip()
    df['Przekroj'] = df['Przekrój '].astype(str).str.strip()

    df['Data'] = pd.to_datetime(df['Data i czas startu zlecenia'], errors='coerce').ffill()

    for c in POMIARY:
        df[c] = pd.to_numeric(df[c], errors='coerce')

    df['LSL'] = pd.to_numeric(df['Wysokość zagniatania LSL w mm'], errors='coerce')
    df['USL'] = pd.to_numeric(df['Wysokośc zaciskania USL w mm'], errors='coerce')
    df['ma_pomiar'] = df[POMIARY].notna().any(axis=1)

    # Tylko interesujące linie
    df = df[df['Linia'].isin(LINIE_SPC)].copy()

    # Przypisanie bloków (connector=1, plug=2) per zlecenie
    # Linie bez struktury CP (np. HPC Uncooled) dostają blok=1
    wyniki = []
    for (z, linia), grp in df.groupby(['Numer zlecenia ', 'Linia'], sort=False):
        grp = grp.reset_index(drop=True)
        cp_rows = grp[grp['Typ'] == 'CP']
        unikalne_cp = list(dict.fromkeys(cp_rows['Numer kontaktu'].tolist()))

        if not unikalne_cp:
            # Brak CP — cała linia to jeden blok (np. HPC z samym PE/DC)
            grp['blok'] = 1
        else:
            cp_to_blok = {cp: i + 1 for i, cp in enumerate(unikalne_cp)}
            current_blok = None
            bloki = []
            for _, row in grp.iterrows():
                if row['Typ'] == 'CP':
                    current_blok = cp_to_blok.get(row['Numer kontaktu'])
                bloki.append(current_blok)
            grp['blok'] = bloki

        wyniki.append(grp)

    df_b = pd.concat(wyniki, ignore_index=True) if wyniki else pd.DataFrame()

    # Budowanie podgrup Xbar-R
    # Podgrupa = jedno zlecenie + linia + blok + typ + maszyna + przekroj
    df_pom = df_b[df_b['ma_pomiar']].copy()

    podgrupy = []
    grp_cols = ['Numer zlecenia ', 'Linia', 'blok', 'Typ', 'Maszyna', 'Przekroj']

    for keys, grp in df_pom.groupby(grp_cols, sort=False):
        z, linia, blok, typ, masz, przekroj = keys
        vals = grp[POMIARY].values.flatten()
        vals = vals[~np.isnan(vals)]
        if len(vals) == 0:
            continue

        n = len(vals)
        lsl = grp['LSL'].dropna().iloc[0] if grp['LSL'].notna().any() else np.nan
        usl = grp['USL'].dropna().iloc[0] if grp['USL'].notna().any() else np.nan
        data = grp['Data'].iloc[0]

        podgrupy.append({
            'Zlecenie': z,
            'Linia': linia,
            'Blok': blok,
            'Typ': typ,
            'Maszyna': masz,
            'Przekroj': przekroj,
            'Data': data,
            'Tydzien': int(data.isocalendar().week),
            'Rok': int(data.year),
            'Xbar': float(np.mean(vals)),
            'R': float(np.max(vals) - np.min(vals)),
            'n': n,
            'LSL': float(lsl) if not np.isnan(lsl) else np.nan,
            'USL': float(usl) if not np.isnan(usl) else np.nan,
        })

    return pd.DataFrame(podgrupy)


# ─── OBLICZENIA SPC ──────────────────────────────────────────────────────────

def oblicz_granice(df_pg: pd.DataFrame) -> dict:
    """Oblicza granice kontrolne Xbar-R dla zestawu podgrup."""
    if df_pg.empty:
        return {}

    # Używamy stałych dla n=3 (dominujący rozmiar podgrupy)
    n_typowy = int(df_pg['n'].mode().iloc[0]) if not df_pg.empty else 3
    n_typowy = max(2, min(n_typowy, 5))  # clamp do obsługiwanych
    st = SPC_STALE[n_typowy]

    Xbarbar = df_pg['Xbar'].mean()
    Rbar = df_pg['R'].mean()

    UCL_x = Xbarbar + st['A2'] * Rbar
    LCL_x = Xbarbar - st['A2'] * Rbar
    UCL_r = st['D4'] * Rbar
    LCL_r = st['D3'] * Rbar

    # Cp / Cpk
    lsl = df_pg['LSL'].dropna().mean()
    usl = df_pg['USL'].dropna().mean()

    sigma_hat = Rbar / st['d2'] if st['d2'] > 0 else np.nan
    cp = cpk = np.nan
    if not (np.isnan(lsl) or np.isnan(usl) or np.isnan(sigma_hat) or sigma_hat == 0):
        cp = (usl - lsl) / (6 * sigma_hat)
        cpu = (usl - Xbarbar) / (3 * sigma_hat)
        cpl = (Xbarbar - lsl) / (3 * sigma_hat)
        cpk = min(cpu, cpl)

    return {
        'Xbarbar': Xbarbar,
        'Rbar': Rbar,
        'UCL_x': UCL_x,
        'LCL_x': LCL_x,
        'UCL_r': UCL_r,
        'LCL_r': LCL_r,
        'LSL': lsl,
        'USL': usl,
        'sigma_hat': sigma_hat,
        'Cp': cp,
        'Cpk': cpk,
        'n_podgrup': len(df_pg),
    }


def wykryj_sygnaly(df_pg: pd.DataFrame, granice: dict) -> pd.Series:
    """
    Reguły Nelson/Western Electric:
    1. Punkt poza UCL/LCL (3σ)
    2. 2 z 3 kolejnych poza 2σ po tej samej stronie
    3. 4 z 5 kolejnych poza 1σ po tej samej stronie
    4. 8 kolejnych po tej samej stronie linii centralnej
    """
    if df_pg.empty or not granice:
        return pd.Series([], dtype=str)

    x = df_pg['Xbar'].values
    cl = granice['Xbarbar']
    sigma3 = granice['UCL_x'] - cl
    sigma1 = sigma3 / 3
    sigma2 = 2 * sigma1

    sygnaly = np.full(len(x), '', dtype=object)

    for i in range(len(x)):
        # Reguła 1: poza 3σ
        if abs(x[i] - cl) > sigma3:
            sygnaly[i] = 'poza_3sigma'
            continue
        # Reguła 2: 2 z 3 poza 2σ
        if i >= 2:
            seg = x[i-2:i+1]
            if sum(v > cl + sigma2 for v in seg) >= 2 or sum(v < cl - sigma2 for v in seg) >= 2:
                sygnaly[i] = 'reguła_2z3'
                continue
        # Reguła 3: 4 z 5 poza 1σ
        if i >= 4:
            seg = x[i-4:i+1]
            if sum(v > cl + sigma1 for v in seg) >= 4 or sum(v < cl - sigma1 for v in seg) >= 4:
                sygnaly[i] = 'reguła_4z5'
                continue
        # Reguła 4: 8 z jednej strony
        if i >= 7:
            seg = x[i-7:i+1]
            if all(v > cl for v in seg) or all(v < cl for v in seg):
                sygnaly[i] = 'reguła_8'

    return pd.Series(sygnaly, index=df_pg.index)


# ─── RYSOWANIE KART ──────────────────────────────────────────────────────────

def rysuj_karte_xbar_r(df_pg: pd.DataFrame, granice: dict, tytul: str) -> go.Figure:
    if df_pg.empty:
        return go.Figure()

    df_pg = df_pg.sort_values('Data').reset_index(drop=True)
    sygnaly = wykryj_sygnaly(df_pg, granice)
    df_pg['sygnal'] = sygnaly.values

    # Etykiety na osi X: numer zlecenia skrócony + data
    labels = [f"{row['Data'].strftime('%d.%m')}\n{row['Zlecenie'][-4:]}"
              for _, row in df_pg.iterrows()]

    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=('X̄ — Średnia podgrupy', 'R — Rozstęp podgrupy'),
        vertical_spacing=0.12,
        row_heights=[0.6, 0.4],
    )

    # ── Karta Xbar ──
    # Strefa wypełnienia ±1σ, ±2σ, ±3σ
    x_idx = list(range(len(df_pg)))
    cl = granice['Xbarbar']
    sigma3 = granice['UCL_x'] - cl
    sigma1 = sigma3 / 3
    sigma2 = 2 * sigma1

    for mult, alpha in [(3, 0.04), (2, 0.07), (1, 0.11)]:
        fig.add_trace(go.Scatter(
            x=x_idx + x_idx[::-1],
            y=[cl + mult * sigma1] * len(x_idx) + [cl - mult * sigma1] * len(x_idx),
            fill='toself',
            fillcolor=f'rgba(37,99,235,{alpha})',
            line=dict(width=0),
            showlegend=False,
            hoverinfo='skip',
        ), row=1, col=1)

    # UCL / LCL / CL
    for val, name, color, dash in [
        (granice['UCL_x'], 'UCL', COL_UCL, 'dash'),
        (granice['LCL_x'], 'LCL', COL_LCL, 'dash'),
        (cl, 'X̄̄', COL_CL, 'solid'),
    ]:
        fig.add_hline(y=val, line_color=color, line_dash=dash, line_width=1.5,
                      annotation_text=f'{name}={val:.4f}',
                      annotation_position='right',
                      annotation_font_size=10,
                      row=1, col=1)

    # USL / LSL
    if not np.isnan(granice.get('USL', np.nan)):
        fig.add_hline(y=granice['USL'], line_color=COL_USL, line_dash='dot', line_width=2,
                      annotation_text=f"USL={granice['USL']:.3f}",
                      annotation_position='right',
                      annotation_font_size=10,
                      row=1, col=1)
    if not np.isnan(granice.get('LSL', np.nan)):
        fig.add_hline(y=granice['LSL'], line_color=COL_LSL, line_dash='dot', line_width=2,
                      annotation_text=f"LSL={granice['LSL']:.3f}",
                      annotation_position='right',
                      annotation_font_size=10,
                      row=1, col=1)

    # Punkty — normalne
    mask_ok = df_pg['sygnal'] == ''
    mask_warn = df_pg['sygnal'].isin(['reguła_2z3', 'reguła_4z5', 'reguła_8'])
    mask_out = df_pg['sygnal'] == 'poza_3sigma'

    hover_xbar = [
        f"Zlecenie: {row['Zlecenie']}<br>"
        f"Data: {row['Data'].strftime('%d.%m.%Y %H:%M')}<br>"
        f"X̄ = {row['Xbar']:.4f}<br>"
        f"R = {row['R']:.4f}<br>"
        f"n = {row['n']}"
        for _, row in df_pg.iterrows()
    ]

    # Linia łącząca
    fig.add_trace(go.Scatter(
        x=x_idx, y=df_pg['Xbar'],
        mode='lines',
        line=dict(color=COL_XBAR, width=1.5),
        showlegend=False,
        hoverinfo='skip',
    ), row=1, col=1)

    # Punkty OK
    if mask_ok.any():
        fig.add_trace(go.Scatter(
            x=[i for i, m in enumerate(mask_ok) if m],
            y=df_pg.loc[mask_ok, 'Xbar'],
            mode='markers',
            marker=dict(color=COL_XBAR, size=8, symbol='circle'),
            name='OK',
            customdata=[hover_xbar[i] for i, m in enumerate(mask_ok) if m],
            hovertemplate='%{customdata}<extra></extra>',
        ), row=1, col=1)

    # Ostrzeżenia
    if mask_warn.any():
        fig.add_trace(go.Scatter(
            x=[i for i, m in enumerate(mask_warn) if m],
            y=df_pg.loc[mask_warn, 'Xbar'],
            mode='markers',
            marker=dict(color=COL_WARN, size=10, symbol='diamond',
                        line=dict(color='white', width=1)),
            name='Ostrzeżenie',
            customdata=[hover_xbar[i] for i, m in enumerate(mask_warn) if m],
            hovertemplate='%{customdata}<extra></extra>',
        ), row=1, col=1)

    # Poza granicami
    if mask_out.any():
        fig.add_trace(go.Scatter(
            x=[i for i, m in enumerate(mask_out) if m],
            y=df_pg.loc[mask_out, 'Xbar'],
            mode='markers',
            marker=dict(color=COL_OUT, size=12, symbol='x',
                        line=dict(color=COL_OUT, width=2)),
            name='Poza kontrolą',
            customdata=[hover_xbar[i] for i, m in enumerate(mask_out) if m],
            hovertemplate='%{customdata}<extra></extra>',
        ), row=1, col=1)

    # ── Karta R ──
    rbar = granice['Rbar']
    ucl_r = granice['UCL_r']

    fig.add_hline(y=ucl_r, line_color=COL_UCL, line_dash='dash', line_width=1.5,
                  annotation_text=f'UCL={ucl_r:.4f}',
                  annotation_position='right',
                  annotation_font_size=10,
                  row=2, col=1)
    fig.add_hline(y=rbar, line_color=COL_CL, line_dash='solid', line_width=1.5,
                  annotation_text=f'R̄={rbar:.4f}',
                  annotation_position='right',
                  annotation_font_size=10,
                  row=2, col=1)

    mask_r_out = df_pg['R'] > ucl_r

    fig.add_trace(go.Bar(
        x=x_idx, y=df_pg['R'],
        marker_color=[COL_OUT if m else COL_R for m in mask_r_out],
        marker_line_width=0,
        name='R',
        hovertemplate='R = %{y:.4f}<extra></extra>',
    ), row=2, col=1)

    # Separator tygodni — pionowe linie
    if 'Tydzien' in df_pg.columns:
        prev_week = None
        for i, (_, row) in enumerate(df_pg.iterrows()):
            if prev_week is not None and row['Tydzien'] != prev_week:
                for r in [1, 2]:
                    fig.add_vline(
                        x=i - 0.5,
                        line_color='rgba(100,100,100,0.3)',
                        line_dash='dot',
                        line_width=1,
                        row=r, col=1,
                    )
                # Adnotacja tygodnia
                fig.add_annotation(
                    x=i - 0.5, y=1.02,
                    yref='paper',
                    text=f'Tydz.{row["Tydzien"]}',
                    showarrow=False,
                    font=dict(size=9, color='#6b7280'),
                    xanchor='center',
                )
            prev_week = row['Tydzien']

    # Oś X — etykiety
    tickvals = list(range(0, len(df_pg), max(1, len(df_pg) // 20)))
    ticktext = [labels[i] for i in tickvals]

    fig.update_xaxes(
        tickvals=tickvals, ticktext=ticktext,
        tickfont=dict(size=8),
        row=1, col=1,
    )
    fig.update_xaxes(
        tickvals=tickvals, ticktext=ticktext,
        tickfont=dict(size=8),
        row=2, col=1,
    )

    fig.update_layout(
        title=dict(text=tytul, font=dict(size=14, color='#1e3a5f')),
        height=550,
        margin=dict(l=60, r=120, t=60, b=40),
        plot_bgcolor='#f8fafc',
        paper_bgcolor='white',
        legend=dict(
            orientation='h', y=-0.08, x=0.5, xanchor='center',
            font=dict(size=10),
        ),
        hovermode='x unified',
    )

    fig.update_yaxes(gridcolor='#e2e8f0', gridwidth=1)

    return fig


# ─── WIDŻETY METRYKI Cp/Cpk ──────────────────────────────────────────────────

def pokaz_metryki_spc(granice: dict):
    cp = granice.get('Cp', np.nan)
    cpk = granice.get('Cpk', np.nan)
    n = granice.get('n_podgrup', 0)

    def kolor_cpk(v):
        if np.isnan(v):
            return '#6b7280'
        if v >= 1.67:
            return '#059669'
        if v >= 1.33:
            return '#10b981'
        if v >= 1.00:
            return '#f59e0b'
        return '#dc2626'

    def status_cpk(v):
        if np.isnan(v):
            return '—'
        if v >= 1.67:
            return '✅ Doskonały'
        if v >= 1.33:
            return '✅ Zdolny'
        if v >= 1.00:
            return '⚠️ Marginalny'
        return '❌ Niezdolny'

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(f"""
        <div style="background:#f0fdf4;border-left:4px solid #059669;
                    padding:12px 16px;border-radius:8px;margin-bottom:8px;">
            <div style="font-size:0.75rem;color:#6b7280;text-transform:uppercase;
                        letter-spacing:0.05em;margin-bottom:4px;">Cp</div>
            <div style="font-size:2rem;font-weight:700;color:#1e3a5f;line-height:1;">
                {f'{cp:.3f}' if not np.isnan(cp) else '—'}
            </div>
            <div style="font-size:0.75rem;color:#6b7280;margin-top:4px;">
                Zdolność potencjalna
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        c = kolor_cpk(cpk if not np.isnan(cpk) else np.nan)
        st.markdown(f"""
        <div style="background:#fefce8;border-left:4px solid {c};
                    padding:12px 16px;border-radius:8px;margin-bottom:8px;">
            <div style="font-size:0.75rem;color:#6b7280;text-transform:uppercase;
                        letter-spacing:0.05em;margin-bottom:4px;">Cpk</div>
            <div style="font-size:2rem;font-weight:700;color:{c};line-height:1;">
                {f'{cpk:.3f}' if not np.isnan(cpk) else '—'}
            </div>
            <div style="font-size:0.75rem;color:#6b7280;margin-top:4px;">
                {status_cpk(cpk)}
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div style="background:#eff6ff;border-left:4px solid #2563eb;
                    padding:12px 16px;border-radius:8px;margin-bottom:8px;">
            <div style="font-size:0.75rem;color:#6b7280;text-transform:uppercase;
                        letter-spacing:0.05em;margin-bottom:4px;">Podgrupy</div>
            <div style="font-size:2rem;font-weight:700;color:#1e3a5f;line-height:1;">
                {n}
            </div>
            <div style="font-size:0.75rem;color:#6b7280;margin-top:4px;">
                X̄̄ = {granice.get('Xbarbar', 0):.4f} mm
            </div>
        </div>
        """, unsafe_allow_html=True)


# ─── GŁÓWNY WIDOK SPC ────────────────────────────────────────────────────────

def pokaz_spc(plik_excel):
    st.subheader('📈 Karty kontrolne SPC — Xbar-R')

    with st.spinner('Przygotowuję dane SPC...'):
        df_pg = przygotuj_dane_spc(plik_excel)

    if df_pg.empty:
        st.warning('Brak danych pomiarowych dla wybranych linii.')
        return

    # ── FILTRY ──
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)

    with col_f1:
        dostepne_linie = sorted(df_pg['Linia'].unique(),
                                key=lambda x: list(LINIE_SPC).index(x) if x in LINIE_SPC else 99)
        wybrana_linia = st.selectbox(
            'Linia',
            dostepne_linie,
            format_func=lambda x: NAZWY_LINII.get(x, x),
        )

    df_linia = df_pg[df_pg['Linia'] == wybrana_linia]

    with col_f2:
        bloki = sorted(df_linia['Blok'].dropna().unique())
        if len(bloki) > 1:
            wybrany_blok = st.selectbox(
                'Strona kabla',
                bloki,
                format_func=lambda x: BLOK_NAZWA.get(int(x), f'Blok {x}') if x else '—',
            )
        else:
            wybrany_blok = bloki[0] if bloki else None
            st.selectbox('Strona kabla',
                         [BLOK_NAZWA.get(int(wybrany_blok), '—') if wybrany_blok else '—'],
                         disabled=True)

    df_blok = df_linia[df_linia['Blok'] == wybrany_blok] if wybrany_blok else df_linia

    with col_f3:
        typy = sorted(df_blok['Typ'].unique())
        wybrany_typ = st.selectbox('Typ przewodu', typy)

    df_typ = df_blok[df_blok['Typ'] == wybrany_typ]

    with col_f4:
        maszyny = sorted(df_typ['Maszyna'].unique())
        wybrana_maszyna = st.selectbox('Maszyna', maszyny)

    df_final = df_typ[df_typ['Maszyna'] == wybrana_maszyna].copy()

    # Filtr tygodni
    dostepne_tygodnie = sorted(df_final['Tydzien'].unique())
    if len(dostepne_tygodnie) > 1:
        zakres = st.select_slider(
            'Zakres tygodni (numer tygodnia ISO)',
            options=dostepne_tygodnie,
            value=(dostepne_tygodnie[0], dostepne_tygodnie[-1]),
        )
        df_final = df_final[df_final['Tydzien'].between(zakres[0], zakres[1])]

    if df_final.empty:
        st.info('Brak danych dla wybranej kombinacji filtrów.')
        return

    # ── OBLICZENIA ──
    granice = oblicz_granice(df_final)

    # ── METRYKI ──
    blok_str = BLOK_NAZWA.get(int(wybrany_blok), '') if wybrany_blok else ''
    linia_str = NAZWY_LINII.get(wybrana_linia, wybrana_linia)
    tytul_karty = f'{linia_str} | {blok_str} | {wybrany_typ} | {wybrana_maszyna}'

    pokaz_metryki_spc(granice)

    st.markdown('<br>', unsafe_allow_html=True)

    # Legenda sygnałów
    with st.expander('ℹ️ Legenda reguł sygnałowych', expanded=False):
        st.markdown("""
        | Symbol | Reguła | Interpretacja |
        |--------|--------|---------------|
        | 🔴 ✕ | Poza 3σ | Punkt poza granicami kontrolnymi UCL/LCL |
        | 🟠 ◆ | 2 z 3 poza 2σ | 2 z 3 kolejnych punktów po tej samej stronie za 2σ |
        | 🟠 ◆ | 4 z 5 poza 1σ | 4 z 5 kolejnych punktów po tej samej stronie za 1σ |
        | 🟠 ◆ | 8 po jednej stronie | 8 kolejnych punktów po tej samej stronie X̄̄ |
        | 🟡 — | USL/LSL | Granice specyfikacji (z arkusza) |
        | ⬜ strefy | ±1σ, ±2σ, ±3σ | Strefy kontrolne (odcienie niebieskiego) |
        """)

    # ── KARTA ──
    fig = rysuj_karte_xbar_r(df_final, granice, tytul_karty)
    st.plotly_chart(fig, use_container_width=True)

    # ── TABELA DANYCH ──
    with st.expander('📋 Dane podgrup', expanded=False):
        df_show = df_final[['Data', 'Zlecenie', 'Tydzien', 'Xbar', 'R', 'n', 'LSL', 'USL']].copy()
        df_show['Data'] = df_show['Data'].dt.strftime('%d.%m.%Y %H:%M')
        df_show = df_show.rename(columns={
            'Data': 'Data', 'Zlecenie': 'Nr zlecenia', 'Tydzien': 'Tydz.',
            'Xbar': 'X̄ (mm)', 'R': 'R (mm)', 'n': 'n próbek',
        })
        df_show['X̄ (mm)'] = df_show['X̄ (mm)'].round(4)
        df_show['R (mm)'] = df_show['R (mm)'].round(4)
        st.dataframe(df_show.reset_index(drop=True), use_container_width=True, hide_index=True)