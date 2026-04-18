import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom # Fallback dla Pythona < 3.9
import io
import zipfile
import re
import unicodedata
import hashlib
import time
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

# --- KONFIGURACJA BIZNESOWA ---
MAX_ZAD_LEN = 23
MAX_OPIS_LEN = 1500
WARN_AMOUNT_THRESHOLD = 1_000_000_000      # 1 miliard złotych
ERROR_AMOUNT_THRESHOLD = 100_000_000_000   # 100 miliardów złotych

# Globalna kompilacja regexu XML
XML_CLEAN_RE = re.compile(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]')

# --- FUNKCJE POMOCNICZE ---

def sanitize_xml(text, context=None, stats=None):
    if pd.isna(text) or text is None:
        return ""
    orig_text = str(text)
    orig_len = len(orig_text)
    cleaned_text = XML_CLEAN_RE.sub('', orig_text)
    if stats is not None and orig_len != len(cleaned_text):
        stats['sanitized_chars'] += (orig_len - len(cleaned_text))
        if context: stats['sanitized_details'].add(context)
    return cleaned_text

def clean_id(value, length=None, strict_mode=True):
    if pd.isna(value): return None
    val_str = str(value).strip().replace('§', '')
    if ',' in val_str or val_str.count('.') > 1: 
        if strict_mode: return None
        else: val_str = val_str.replace('.', '').replace(',', '')
    val_str = re.sub(r'\.0+$', '', val_str)
    try:
        val_float = float(val_str)
        if val_float.is_integer():
            val_int = int(val_float)
            return str(val_int).zfill(length) if length else str(val_int)
        return None
    except ValueError:
        if val_str.isdigit(): return val_str.zfill(length) if length else val_str
        return None

def parse_kwota(val, strict_mode=True):
    if pd.isna(val) or val == "": return None
    def to_grosze(numeric_val):
        try:
            dec = Decimal(str(numeric_val))
            return int(dec.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) * 100)
        except InvalidOperation: return None

    if isinstance(val, (int, float)): return to_grosze(val)
    v = str(val).strip()
    if not v: return None
    try:
        v = v.replace('–', '-').replace('—', '-').replace('−', '-')
        v = re.sub(r'[\s\u202f\u2009\xa0]+', '', v)
        if v.startswith('(') and v.endswith(')'): v = '-' + v[1:-1]
        if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
        elif ',' in v:
            parts = v.split(',')
            if len(parts) == 2 and len(parts[1]) in [1, 2]: v = v.replace(',', '.')
            elif len(parts) == 2 and len(parts[1]) == 3:
                if strict_mode: return None 
                else: v = v.replace(',', '')
            else: v = v.replace(',', '')
        elif '.' in v and v.count('.') > 1: v = v.replace('.', '')
        return to_grosze(v)
    except Exception: return None

def normalize_filename(name):
    name_str = str(name).upper()
    nfkd_form = unicodedata.normalize('NFKD', name_str)
    name_ascii = nfkd_form.encode('ASCII', 'ignore').decode('utf-8')
    name_clean = re.sub(r'[^A-Z0-9]', '_', name_ascii)
    name_clean = re.sub(r'_+', '_', name_clean).strip('_')
    hash_val = hashlib.sha1(str(name).encode()).hexdigest()[:6]
    if name_clean: return f"{name_clean[:60]}_{hash_val}"
    return f"BRAK_NAZWY_{hash_val}"

def load_mapping_dict(uploaded_file):
    try:
        df_dict = pd.read_excel(uploaded_file, sheet_name='Słowniki')
        df_dict.columns = df_dict.columns.str.strip()
        if 'Nazwa_Excel' in df_dict.columns and 'Nazwa_Systemowa' in df_dict.columns:
            keys = df_dict['Nazwa_Excel'].astype(str).str.strip()
            vals = df_dict['Nazwa_Systemowa'].astype(str).str.strip()
            return {k: v for k, v in dict(zip(keys, vals)).items() if k != 'nan'}
    except ValueError: pass 
    except Exception as e: st.warning(f"⚠️ Ostrzeżenie ze Słownika: ({e})")
    return {}

def format_pln(amount):
    return f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")

def create_xml(data_frame, doc_params, unit_name, mapping_dict, typ_str, stats, typ_zmiany_val):
    if data_frame.empty: return ""
        
    root = ET.Element("Plan", wersja="1.0")
    typ_xml_node = ET.SubElement(root, typ_str)
    
    if 'Uzasadnienie' in data_frame.columns:
        uzas_list = data_frame['Uzasadnienie'].astype(str).str.strip()
        valid_uzas = [u for u in uzas_list.unique() if u.lower() not in ['nan', 'none', '']]
        uzasadnienie_raw = " | ".join(valid_uzas) if valid_uzas else doc_params['opis']
    else:
        uzasadnienie_raw = doc_params['opis']

    finalny_opis = sanitize_xml(uzasadnienie_raw, f"Uzasadnienie ({unit_name})", stats)[:MAX_OPIS_LEN]
    dysponent_sys = mapping_dict.get(unit_name, unit_name)
    bezpieczny_dysponent = sanitize_xml(dysponent_sys, f"Dysponent jednostki", stats)

    dok_node = ET.SubElement(typ_xml_node, "Dokument", 
                             PodstawaPrawna="DP", TYP="2" if typ_str == "Wydatki" else "1", 
                             NR_DOK=sanitize_xml(doc_params['nr_dok'], "Nr Dokumentu", stats), 
                             DATA_DOK=doc_params['data_dok'], 
                             ROK=doc_params['rok'], MC=doc_params['mc'], 
                             ROK_BUD=doc_params['rok'], ROK_KSIEGOWY=doc_params['rok'], 
                             MC_KSIEG=doc_params['mc'], OPIS=finalny_opis, 
                             P_PIERWOTNY="N", P_WNW="N", TYP_ZMIANY=typ_zmiany_val)
    
    df_sorted = data_frame.copy()
    stats['audit_before'] += len(df_sorted)
    
    if 'Fundusz' in df_sorted.columns:
        df_sorted['Sposob_finansowania'] = df_sorted['Fundusz'].fillna('WG')
    elif 'Sposób finansowania' in df_sorted.columns:
        df_sorted['Sposob_finansowania'] = df_sorted['Sposób finansowania'].fillna('WG')
    else:
        df_sorted['Sposob_finansowania'] = 'WG'
    
    df_sorted['Zad_Sys'] = df_sorted['Zadanie'].astype(str).str.strip().apply(lambda x: mapping_dict.get(x, x))
    df_sorted.loc[df_sorted['Zad_Sys'].isin(['nan', '', 'None']), 'Zad_Sys'] = "000000000"
    df_sorted['Zad_Sys'] = df_sorted['Zad_Sys'].apply(lambda x: sanitize_xml(str(x)[:MAX_ZAD_LEN], "Zadanie", stats))
    
    group_cols = ['Dzial_clean', 'Rozdzial_clean', 'Paragraf_clean', 'Sposob_finansowania', 'Zad_Sys']
    
    before_drop = len(df_sorted)
    df_sorted = df_sorted.dropna(subset=group_cols)
    dropped_count = before_drop - len(df_sorted)
    if dropped_count > 0: stats['dropped_na'] += dropped_count
    
    group_sizes = df_sorted.groupby(group_cols).size()
    merged_groups = group_sizes[group_sizes > 1]
    for name, count in merged_groups.items():
        dz, ro, pa, sf, zs = name
        stats['merged_details'].append(f"{unit_name} | Dz:{dz} Rozdz:{ro} Par:{pa} -> skompresowano {count} wiersze do salda.")

    df_grouped = df_sorted.groupby(group_cols, as_index=False)['Zmiana_num'].sum()
    stats['audit_after'] += len(df_grouped)

    df_grouped['Dzial_num'] = pd.to_numeric(df_grouped['Dzial_clean'], errors='coerce')
    df_grouped['Rozdzial_num'] = pd.to_numeric(df_grouped['Rozdzial_clean'], errors='coerce')
    df_grouped['Paragraf_num'] = pd.to_numeric(df_grouped['Paragraf_clean'], errors='coerce')
    
    df_grouped = df_grouped.sort_values(by=['Dzial_num', 'Rozdzial_num', 'Paragraf_num'], na_position='last', kind='mergesort')
    
    lp = 1
    for row in df_grouped.itertuples(index=False):
        dz, ro, pa = row.Dzial_clean, row.Rozdzial_clean, row.Paragraf_clean
        kwota_grosze = getattr(row, 'Zmiana_num', None)
        zad_sys = row.Zad_Sys
        sposob_fin = str(row.Sposob_finansowania).strip()
        
        if not dz or not ro or not pa or pd.isna(kwota_grosze):
            stats['runtime_errors_count'] += 1
            if len(stats['runtime_errors_list']) < 100:
                stats['runtime_errors_list'].append(f"Dz:{dz}, Rozdz:{ro}, Par:{pa}, Kwota Błąd")
            continue
            
        if kwota_grosze == 0: 
            stats['skipped_zeros'] += 1
            continue
            
        kwota_zl = Decimal(int(kwota_grosze)) / 100
            
        if abs(kwota_zl) > ERROR_AMOUNT_THRESHOLD:
            raise ValueError(f"KRYTYCZNY BŁĄD BIZNESOWY: Kwota {format_pln(kwota_zl)} zł (Dz:{dz} R:{ro} P:{pa}) w {unit_name} przekracza absolutny limit bezpieczeństwa.")
            
        if abs(kwota_zl) > WARN_AMOUNT_THRESHOLD:
            stats['suspicious_amounts'] += 1
            stats['suspicious_list'].append(f"{unit_name} | Dz:{dz} Rozdz:{ro} Par:{pa} -> **{format_pln(kwota_zl)} zł**")

        ET.SubElement(dok_node, "Pozycja", 
                      Dysponent=bezpieczny_dysponent, SposobFinansowania=sanitize_xml(sposob_fin, "Sposób finansowania", stats), 
                      Dzial=dz, Rozdzial=ro, Paragraf=pa, 
                      Pozycja="", Zadanie=zad_sys, 
                      Data=doc_params['data_dok'], Lp=str(lp), Plan=f"{kwota_zl:.2f}")
        lp += 1
    
    if len(dok_node) == 0: return ""
            
    if hasattr(ET, "indent"):
        ET.indent(root, space="  ", level=0)
        xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        xml_str = xml_bytes.decode('utf-8')
    else:
        xml_bytes = ET.tostring(root, encoding='utf-8')
        xml_str = minidom.parseString(xml_bytes).toprettyxml(indent="  ", encoding="utf-8").decode("utf-8")
        
    return re.sub(r'<\?xml[^>]+\?>', '<?xml version="1.0" encoding="UTF-8"?>', xml_str)

# --- STREAMLIT UI ---
st.set_page_config(page_title="Konwerter Budżetowy Rekord", layout="wide")

st.sidebar.header("📝 Dane Dokumentu")
d_date = st.sidebar.date_input("Data dokumentu", datetime.today())
d_nr = st.sidebar.text_input("Numer", "ZMIANA/2026/01")
d_opis = st.sidebar.text_input("Uzasadnienie (fallback)", "Zmiana planu")

st.sidebar.header("⚙️ Ustawienia Księgowe")

# --- UX WIDGET: Jasne opisy Typu Zmiany ---
opcje_typu_zmiany = {
    "10 (domyślnie) - uchwała/zarządzenie": "10",
    "0 - plan początkowy": "0"
}
wybrana_etykieta_zmiany = st.sidebar.selectbox(
    "Typ Zmiany XML", 
    options=list(opcje_typu_zmiany.keys()), 
    index=0, 
    help="Zdecyduj o charakterze prawnym importu do systemu Rekord."
)
typ_zmiany_opcja = opcje_typu_zmiany[wybrana_etykieta_zmiany]
# ------------------------------------------

strict_mode = st.sidebar.checkbox("Restrykcyjna walidacja danych", value=True, help="Zaznaczone: Odrzuca dziwne ułamki i nadmiarowe przecinki. Odznaczone: Próbuje zgadywać i naprawiać.")
audit_mode = st.sidebar.checkbox("Tryb Audytu (Analiza Danych)", value=False, help="Wyświetla szczegółowe informacje m.in. o zsumowanych wierszach na identycznej klasyfikacji.")

d_params = {'nr_dok': d_nr, 'data_dok': d_date.strftime("%Y-%m-%d"), 
            'rok': str(d_date.year), 'mc': str(d_date.month), 'opis': d_opis}

st.title("🚀 Generator XML dla Rekord SI")

f = st.file_uploader("Wgraj Excel (arkusze: Zmiany, Słowniki)", type="xlsx")

if f:
    start_time = time.time()
    mapping = load_mapping_dict(f)
    if mapping: st.sidebar.success(f"Wczytano słownik: {len(mapping)} mapowań.")
    
    try:
        df = pd.read_excel(f, sheet_name='Zmiany')
        df.columns = df.columns.str.strip()
        col_p = '§' if '§' in df.columns else 'Paragraf'
        
        type_map = {
            'dochody': 'Dochody', 'dochód': 'Dochody', 'doch': 'Dochody', 'd': 'Dochody', 'dw': 'Dochody',
            'wydatki': 'Wydatki', 'wydatek': 'Wydatki', 'wyd': 'Wydatki', 'w': 'Wydatki', 'wd': 'Wydatki'
        }
        cleaned_typ = df['Typ D/W'].astype(str).str.lower().str.replace(r'[^a-ząćęłńóśżź]', '', regex=True)
        df['Typ_DW_norm'] = cleaned_typ.map(type_map).fillna("BŁĄD")
        
        df['Dzial_clean'] = df['Dział'].apply(lambda x: clean_id(x, 3, strict_mode))
        df['Rozdzial_clean'] = df['Rozdział'].apply(lambda x: clean_id(x, 5, strict_mode))
        df['Paragraf_clean'] = df[col_p].apply(lambda x: clean_id(x, 4, strict_mode))
        df['Zmiana_num'] = df['Zmiana'].apply(lambda x: parse_kwota(x, strict_mode)).astype(pd.Int64Dtype())
        df['Jednostka_clean'] = df['Jednostka'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        
        errors = []
        df_valid = df[df['Jednostka_clean'].notna() & (df['Jednostka_clean'] != '') & (df['Jednostka_clean'] != 'nan')].copy()
        
        for row in df_valid.itertuples():
            r_num = row.Index + 2
            u = str(getattr(row, 'Jednostka_clean', 'Brak'))
            
            if u.isdigit(): errors.append(f"Wiersz {r_num}: Jednostka musi mieć nazwę tekstową ('{u}').")
                
            dz = str(getattr(row, 'Dzial_clean', ''))
            ro = str(getattr(row, 'Rozdzial_clean', ''))
            pa = str(getattr(row, 'Paragraf_clean', ''))
            kwota_grosze = getattr(row, 'Zmiana_num', pd.NA)
            
            if getattr(row, 'Typ_DW_norm', '') == "BŁĄD": 
                errors.append(f"Wiersz {r_num} ({u}): Nierozpoznany Typ D/W.")
            if not dz or len(dz) != 3 or not dz.isdigit(): 
                errors.append(f"Wiersz {r_num} ({u}): Dział musi mieć 3 cyfry (bez znaków specjalnych).")
            if not ro or len(ro) != 5 or not ro.isdigit(): 
                errors.append(f"Wiersz {r_num} ({u}): Rozdział musi mieć 5 cyfr.")
            if not pa or len(pa) != 4 or not pa.isdigit(): 
                errors.append(f"Wiersz {r_num} ({u}): Paragraf musi mieć dokładnie 4 cyfry.")
                
            if pd.isna(kwota_grosze): 
                errors.append(f"Wiersz {r_num} ({u}): Kwota '{getattr(row, 'Zmiana', '')}' jest nieczytelna (Tryb Strict: {strict_mode}).")
            elif abs(Decimal(int(kwota_grosze)) / 100) > ERROR_AMOUNT_THRESHOLD:
                errors.append(f"Wiersz {r_num} ({u}): BŁĄD KRYTYCZNY! Kwota {format_pln(int(kwota_grosze) / 100.0)} zł przekracza limit.")

        if errors:
            st.error(f"🚨 Znaleziono {len(errors)} błędów blokujących generowanie plików.")
            st.download_button("⬇️ Pobierz pełny raport błędów", "\n".join(errors), "bledy.txt")
            for e in errors[:50]: st.write(f"❌ {e}")
            if len(errors) > 50: st.info(f"...oraz {len(errors) - 50} więcej. Pobierz raport .txt.")
            st.stop()

        bilans_grosze = pd.to_numeric(df_valid['Zmiana_num'], errors='coerce').dropna().astype('int64').sum()
        bilans_zl = Decimal(bilans_grosze) / 100
        
        units = sorted(df_valid['Jednostka_clean'].unique())
        z_buf, used_names = io.BytesIO(), set()
        
        stats = {
            'skipped_zeros': 0, 'runtime_errors_count': 0, 'runtime_errors_list': [], 
            'audit_before': 0, 'audit_after': 0, 'sanitized_chars': 0,
            'sanitized_details': set(), 'suspicious_amounts': 0, 'suspicious_list': [],
            'dropped_na': 0, 'merged_details': []
        }
        preview = ""

        with zipfile.ZipFile(z_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for unit in units:
                u_df = df_valid[df_valid['Jednostka_clean'] == unit]
                for t in sorted(x for x in u_df['Typ_DW_norm'].unique() if x in ['Dochody', 'Wydatki']):
                    sub = u_df[u_df['Typ_DW_norm'] == t].reset_index(drop=True)
                    
                    xml = create_xml(sub, d_params, unit, mapping, t, stats, typ_zmiany_opcja)
                    if not xml: continue 
                    if not preview: preview = xml
                    
                    fname = f"Plan_{normalize_filename(unit)}_{t}.xml"
                    c = 1
                    while fname in used_names:
                        fname = f"Plan_{normalize_filename(unit)}_{t}_{c}.xml"
                        c += 1
                    used_names.add(fname)
                    zf.writestr(fname, xml.encode('utf-8'))

        if not used_names:
            st.warning("⚠️ Brak danych do wygenerowania. System usunął puste pola i wyzerowane kwoty.")
            st.stop()

        st.success(f"Sukces! Wygenerowano {len(used_names)} dokumentów XML dla {len(units)} jednostek. ✅")
        
        if bilans_grosze == 0:
            st.info("⚖️ Bilans zmian (globalny) wynosi **0,00 zł** (Idealnie zbilansowane przesunięcie budżetowe).")
        else:
            st.info(f"📈 Bilans zmian (globalny) wynosi: **{format_pln(bilans_zl)} zł** (Zmiana wielkości budżetu).")

        with st.expander("📊 Wyświetl bilans zmian dla poszczególnych jednostek"):
            bilans_jednostek = pd.to_numeric(df_valid['Zmiana_num'], errors='coerce').groupby(df_valid['Jednostka_clean']).sum()
            for j_name, j_grosze in bilans_jednostek.items():
                if pd.isna(j_grosze): j_grosze = 0
                j_zl = Decimal(int(j_grosze)) / 100
                if j_grosze == 0: st.write(f"- {j_name}: **0,00 zł** 🟢")
                else: st.write(f"- {j_name}: **{format_pln(j_zl)} zł** 🔵")
        
        if stats['suspicious_amounts'] > 0:
            st.error(f"🚨 WYKRYTO ZAGROŻENIE: Znaleziono {stats['suspicious_amounts']} kwot pow. {format_pln(WARN_AMOUNT_THRESHOLD)} zł.")
            with st.expander("Szczegóły podejrzanych kwot"):
                for s_item in stats['suspicious_list'][:10]: st.write(f"- {s_item}")

        if stats['dropped_na'] > 0:
            st.error(f"🚨 ODRZUCONO DANE: Usunięto **{stats['dropped_na']}** wierszy z powodu braku kodów klasyfikacji. Sprawdź plik!")

        if audit_mode:
            if stats['audit_before'] > stats['audit_after']:
                st.info(f"🔍 **Tryb Audytu JST:** Zoptymalizowano liczbę wierszy z **{stats['audit_before']}** (Excel) do **{stats['audit_after']}** (XML). Zsumowano {stats['audit_before'] - stats['audit_after']} powiązanych ze sobą operacji.")
                if stats['merged_details']:
                    with st.expander("Zobacz detale skompresowanych podziałek budżetowych"):
                        for m_item in stats['merged_details'][:15]: st.write(f"- {m_item}")
                        if len(stats['merged_details']) > 15: st.write("*...i więcej.*")
            else:
                st.info("🔍 **Tryb Audytu JST:** Nie wykryto wpisów w tej samej klasyfikacji wymagających sumowania.")

        if stats['sanitized_chars'] > 0:
            st.warning(f"🧹 Usunięto **{stats['sanitized_chars']}** znaków specjalnych niedozwolonych przez standard XML 1.0.")
            with st.expander("Zlokalizowane w:"):
                for detail in list(stats['sanitized_details'])[:5]: st.write(f"- {detail}")

        if stats['skipped_zeros'] > 0: 
            st.info(f"ℹ️ Zoptymalizowano (zsumowano do 0.00 lub pominięto): {stats['skipped_zeros']} pozycji.")
            
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📦 Pobierz paczkę ZIP", z_buf.getvalue(), 
                               f"Eksport_Rekord_{d_date.strftime('%Y%m%d')}.zip", "application/zip", use_container_width=True)
        
        with st.expander("🔍 Podgląd pierwszego wygenerowanego dokumentu XML"): st.code(preview, "xml")
        
        st.caption(f"⏱️ Czas generowania: {round(time.time() - start_time, 2)} s | Wczytano i przetworzono {len(df)} wierszy źródłowych.")
            
    except Exception as e: st.error(f"Błąd krytyczny aplikacji: {e}")