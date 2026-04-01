import pandas as pd
import streamlit as st
import datetime
import io

# --- KONFIGURACJA APLIKACJI ---
st.set_page_config(page_title="Analizator Q-Gate", layout="wide")
st.title("📊 Dashboard analizy próbek Q-Gate")


# --- SILNIK ANALITYCZNY ---
@st.cache_data
def wczytaj_dane(plik_excel):
    # Zmiana: teraz funkcja przyjmuje wgrany plik, a nie ścieżkę tekstową
    df = pd.read_excel(plik_excel, sheet_name="Qgate 2026", header=2, engine='openpyxl')

    df['Numer zlecenia '] = df['Numer zlecenia '].ffill()
    df['NR LINII '] = df['NR LINII '].ffill()
    df['Data i czas startu zlecenia'] = pd.to_datetime(df['Data i czas startu zlecenia'], errors='coerce').ffill()
    df['Data_kalendarzowa'] = df['Data i czas startu zlecenia'].dt.date

    df = df.dropna(subset=['Numer zlecenia '])
    df['NR LINII '] = df['NR LINII '].astype(str).str.strip().str.upper()
    df['Typ pojedyńczego przewodu'] = df['Typ pojedyńczego przewodu'].fillna('').astype(str).str.strip().str.upper()

    kolumny_probki = [
        'Wysokość zaciskania - zrywy',
        'Wysokość zaciskania -  kontrola przekroju na mikrografie',
        'Wysokość zaciskania - próbka poglądowa'
    ]

    df['Wiersz_ma_probke'] = df[kolumny_probki].notna().any(axis=1)
    df['Czy_to_PP'] = df['Typ pojedyńczego przewodu'] == 'PP'

    zlecenia_z_zestawem = df.groupby(['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa'])[
        'Wiersz_ma_probke'].any().reset_index()
    prawidlowe_zestawy = zlecenia_z_zestawem[zlecenia_z_zestawem['Wiersz_ma_probke'] == True]

    df_zestawy = pd.merge(df, prawidlowe_zestawy[['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa']],
                          on=['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa'],
                          how='inner')

    df_do_zliczenia = df_zestawy[(df_zestawy['Wiersz_ma_probke'] == True) | (df_zestawy['Czy_to_PP'] == True)]

    df_unikalne_kontakty = df_do_zliczenia.drop_duplicates(
        subset=['Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa', 'Numer kontaktu', 'Typ pojedyńczego przewodu']
    ).copy()

    df_unikalne_kontakty['Ilosc_fizycznych_probek'] = 3

    zlecenia = df_unikalne_kontakty.groupby([
        'Numer zlecenia ', 'NR LINII ', 'Data_kalendarzowa'
    ]).agg(
        Ilosc_zbrakowanych_kontaktow=('Numer kontaktu', 'size'),
        Data_i_czas_startu_zlecenia=('Data i czas startu zlecenia', 'max')
    ).reset_index()

    zlecenia = zlecenia.rename(columns={'Data_i_czas_startu_zlecenia': 'Data i czas startu zlecenia'})
    zlecenia['Ilosc_zestawow'] = 1
    zlecenia['Ilosc_fizycznych_probek'] = zlecenia['Ilosc_zbrakowanych_kontaktow'] * 3

    return zlecenia, df_unikalne_kontakty


def wygeneruj_excel(df_glowne, df_szczegoly=None, nazwa_zakladki_szczegolow="Szczegóły", df_trzecie=None,
                    nazwa_zakladki_trzecie="Dodatkowe"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_glowne.to_excel(writer, index=False, sheet_name='Podsumowanie Linii')
        if df_szczegoly is not None:
            df_szczegoly.to_excel(writer, index=False, sheet_name=nazwa_zakladki_szczegolow)
        if df_trzecie is not None:
            df_trzecie.to_excel(writer, index=False, sheet_name=nazwa_zakladki_trzecie)
    return output.getvalue()


# --- INTERFEJS UŻYTKOWNIKA ---
st.sidebar.header("📁 Wczytaj dane")
wgrany_plik = st.sidebar.file_uploader("Przeciągnij plik Excel Q-Gate:", type=['xlsx', 'xlsm'])
st.sidebar.markdown("---")

if wgrany_plik is None:
    st.info("👈 Aby rozpocząć pracę, wgraj plik Excel z danymi Q-Gate w panelu po lewej stronie.")
else:
    try:
        # Przekazujemy wgrany plik do naszej funkcji
        df_zlecenia, df_kontakty = wczytaj_dane(wgrany_plik)

        st.sidebar.header("🛠️ Ustawienia raportu")

        tryb_raportu = st.sidebar.radio(
            "Wybierz rodzaj raportu:",
            ["Raport zmianowy (dzienny)", "Raport miesięczny"]
        )
        st.sidebar.markdown("---")

        if tryb_raportu == "Raport zmianowy (dzienny)":
            max_date = df_zlecenia['Data i czas startu zlecenia'].max().date()
            wybrana_data = st.sidebar.date_input("1. Wybierz dzień", value=max_date)

            st.sidebar.markdown("**2. Wybierz zakres godzin**")
            od_godziny = st.sidebar.time_input("Od godziny:", value=datetime.time(6, 0))
            do_godziny = st.sidebar.time_input("Do godziny:", value=datetime.time(14, 0))

            start_dt = pd.to_datetime(f"{wybrana_data} {od_godziny}")
            if do_godziny < od_godziny:
                end_dt = pd.to_datetime(f"{wybrana_data} {do_godziny}") + pd.Timedelta(days=1)
            else:
                end_dt = pd.to_datetime(f"{wybrana_data} {do_godziny}")

            mask_zlecenia = (df_zlecenia['Data i czas startu zlecenia'] >= start_dt) & (
                        df_zlecenia['Data i czas startu zlecenia'] <= end_dt)
            df_filtered = df_zlecenia[mask_zlecenia]

            mask_kontakty = (df_kontakty['Data i czas startu zlecenia'] >= start_dt) & (
                        df_kontakty['Data i czas startu zlecenia'] <= end_dt)
            df_kontakty_filtered = df_kontakty[mask_kontakty]

            tytul_raportu = f"Podsumowanie zmiany: {start_dt.strftime('%d.%m.%Y, %H:%M')} - {end_dt.strftime('%H:%M')}"

        else:
            df_zlecenia['Rok_Miesiac'] = df_zlecenia['Data i czas startu zlecenia'].dt.strftime('%Y-%m')
            df_kontakty['Rok_Miesiac'] = df_kontakty['Data i czas startu zlecenia'].dt.strftime('%Y-%m')

            lista_miesiecy = sorted(df_zlecenia['Rok_Miesiac'].unique(), reverse=True)
            wybrany_miesiac = st.sidebar.selectbox("Wybierz miesiąc do analizy:", lista_miesiecy)

            df_filtered = df_zlecenia[df_zlecenia['Rok_Miesiac'] == wybrany_miesiac]
            df_kontakty_filtered = df_kontakty[df_kontakty['Rok_Miesiac'] == wybrany_miesiac]

            tytul_raportu = f"Podsumowanie miesiąca: {wybrany_miesiac}"

        # --- GŁÓWNY WIDOK ---
        st.subheader(tytul_raportu)

        if df_filtered.empty:
            st.warning("Brak zrealizowanych zestawów z próbkami w wybranym przedziale.")
        else:
            col1, col2, col3 = st.columns(3)
            col1.metric("Suma zbrakowanych kontaktów", f"{df_filtered['Ilosc_zbrakowanych_kontaktow'].sum()} szt.")
            col2.metric("Liczba zrealizowanych zestawów", f"{df_filtered['Ilosc_zestawow'].sum()} szt.")
            col3.metric("🔥 Fizyczne próbki do SAP", f"{df_filtered['Ilosc_fizycznych_probek'].sum()} szt.")

            st.divider()

            if tryb_raportu == "Raport miesięczny":
                lewa, prawa = st.columns(2)
            else:
                lewa = st.container()
                prawa = None

            with lewa:
                st.markdown("**Złomowanie według linii produkcyjnych**")
                podsumowanie_linii = df_filtered.groupby('NR LINII ')[
                    ['Ilosc_zbrakowanych_kontaktow', 'Ilosc_zestawow', 'Ilosc_fizycznych_probek']].sum().reset_index()
                podsumowanie_linii = podsumowanie_linii.sort_values(by='Ilosc_fizycznych_probek', ascending=False)

                podsumowanie_linii = podsumowanie_linii.rename(columns={
                    'NR LINII ': 'Linia',
                    'Ilosc_zbrakowanych_kontaktow': 'Zbrakowane kontakty',
                    'Ilosc_zestawow': 'Zestawy',
                    'Ilosc_fizycznych_probek': 'Sztuki do SAP'
                })

                st.dataframe(podsumowanie_linii, width='stretch', hide_index=True)
                st.bar_chart(podsumowanie_linii.set_index('Linia')['Sztuki do SAP'])

            trend_dzienny = None
            if prawa:
                with prawa:
                    st.markdown("**Trend dzienny złomowania w wybranym miesiącu**")
                    trend_dzienny = df_filtered.groupby('Data_kalendarzowa')[
                        'Ilosc_fizycznych_probek'].sum().reset_index()
                    trend_dzienny['Data_kalendarzowa'] = pd.to_datetime(trend_dzienny['Data_kalendarzowa']).dt.strftime(
                        '%d.%m')

                    st.bar_chart(trend_dzienny.set_index('Data_kalendarzowa')['Ilosc_fizycznych_probek'])

            # --- SEKCJA DLA SAP ---
            st.divider()
            st.markdown("### 📋 Numery kontaktów do zezłomowania w SAP")
            st.info("Gotowa ściągawka z numerami kontaktów (Part Number) i sumą fizycznych sztuk do zezłomowania.")

            podsumowanie_kontaktow = df_kontakty_filtered.groupby('Numer kontaktu')[
                'Ilosc_fizycznych_probek'].sum().reset_index()
            podsumowanie_kontaktow = podsumowanie_kontaktow.sort_values(by='Ilosc_fizycznych_probek', ascending=False)

            podsumowanie_kontaktow = podsumowanie_kontaktow.rename(columns={
                'Numer kontaktu': 'Numer kontaktu (PN)',
                'Ilosc_fizycznych_probek': 'Sztuki do wyrzucenia'
            })

            c1, c2, c3 = st.columns([1, 2, 1])
            with c2:
                st.dataframe(podsumowanie_kontaktow, width='stretch', hide_index=True)

            # --- SEKCJA EKSPORTU DO EXCELA ---
            st.divider()
            st.markdown("### 💾 Eksport danych")

            if tryb_raportu == "Raport zmianowy (dzienny)":
                excel_data = wygeneruj_excel(podsumowanie_linii, podsumowanie_kontaktow, "Ściągawka SAP")
                nazwa_eksportu = f"Raport_Zmianowy_{wybrana_data}.xlsx"
            else:
                excel_data = wygeneruj_excel(podsumowanie_linii, podsumowanie_kontaktow, "Ściągawka SAP", trend_dzienny,
                                             "Trend Dzienny")
                nazwa_eksportu = f"Raport_Miesieczny_{wybrany_miesiac}.xlsx"

            st.download_button(
                label="📥 Pobierz ten raport jako plik Excel (.xlsx)",
                data=excel_data,
                file_name=nazwa_eksportu,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Wystąpił błąd podczas analizy pliku. Upewnij się, że to poprawny plik Q-Gate. Szczegóły: {e}")