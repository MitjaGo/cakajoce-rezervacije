"""
Streamlit aplikacija: Filtriranje rezervacij s statusom "Na čakanju"
======================================================================
Naloži 1-6 XLS izvoznih datotek (PMS sistem, HTML-tabela s pripono .xls),
filtrira vrstice s statusom "Na čakanju", kjer je od stolpca
"Datum nastanka" do izbranega datuma filtracije preteklo N ali več dni
(privzeto 4), združi rezultate vseh datotek v en Excel dokument in
omogoči prenos na računalnik.

Zagon:
    pip install -r requirements.txt
    streamlit run app.py
"""

import re
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Rezervacije - Na čakanju", layout="wide")

LOGO_URL = "https://www.adria-ankaran.si//app/uploads/2025/10/logo-Adria.jpg"

header_left, header_right = st.columns([4, 1])
with header_left:
    st.title("📋 Filtriranje rezervacij s statusom \"Na čakanju\"")
with header_right:
    st.image(LOGO_URL, width=150)

st.markdown(
    """
Naloži od **1 do 6** XLS datotek (izvoz iz PMS sistema). Aplikacija bo:
- prebrala podatke iz vsake datoteke,
- obdržala samo vrstice s statusom **Na čakanju**,
- izračunala, koliko dni je preteklo od stolpca **Datum nastanka** do izbranega
  datuma filtracije,
- prikazala vrstice, kjer je preteklo **N ali več dni** (privzeto 4),
- prikazala stolpce: **Code, PMS koda, Objekt, Datum nastanka, Prihod,
  Lastnik rezervacije, Status**,
- združila rezultate vseh naloženih datotek v en Excel dokument, ki ga
  prenesete na svoj računalnik.
"""
)

# ---------------------------------------------------------------------------
# Nastavitve filtra
# ---------------------------------------------------------------------------
col1, col2, col3 = st.columns(3)
with col1:
    filter_date = st.date_input("Datum filtracije", value=date.today())
with col2:
    min_days = st.number_input(
        "Min. dni od 'Datum nastanka' (dolgo čakanje)", min_value=0, value=4, step=1
    )
with col3:
    urgent_days = st.number_input(
        "Maks. dni med nastankom in prihodom (prihod kmalu)",
        min_value=0,
        value=3,
        step=1,
    )

st.caption(
    "Vrstica se prikaže, če je status 'Na čakanju' IN (od nastanka je "
    "preteklo ≥ zgornji prag DNI, ALI je bil prihod že ob rezervaciji "
    "napovedan v roku ≤ zgornji prag PRIHOD - gost mora plačati vnaprej, "
    "zato je treba te rezervacije nujno preveriti)."
)

uploaded_files = st.file_uploader(
    "Naloži XLS datoteke (1-6 datotek)",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
)

REQUIRED_COLS = [
    "Status",
    "PMS koda",
    "Code",
    "Objekt",
    "Datum nastanka",
    "Prihod",
    "Lastnik rezervacije",
]


# ---------------------------------------------------------------------------
# Pomožne funkcije
# ---------------------------------------------------------------------------
def _pick_best_table(content: bytes):
    """PMS izvoz ima v prvi vrstici naslovni <td colspan=...> (npr. 'Premium
    mobile homes'), zato mora biti pravi header v vrstici 1, ne 0. Preizkusi
    header=0 in header=1 ter obdrži različico, ki vsebuje pričakovane
    stolpce (ali ima v vsakem primeru največ prepoznanih stolpcev)."""
    best_df = None
    best_score = -1
    for header_row in (1, 0, None):
        try:
            tables = pd.read_html(BytesIO(content), header=header_row)
        except Exception:
            continue
        if not tables:
            continue
        candidate = max(tables, key=lambda t: t.shape[1])
        candidate.columns = [str(c).strip() for c in candidate.columns]
        score = sum(1 for c in REQUIRED_COLS if c in candidate.columns)
        if score > best_score:
            best_score = score
            best_df = candidate
        if score == len(REQUIRED_COLS):
            break
    return best_df


def parse_file(file) -> "pd.DataFrame | None":
    """Datoteke so v resnici HTML tabela shranjena s pripono .xls."""
    content = file.read()
    df = _pick_best_table(content)
    if df is None:
        # poskusi kot pravi binarni/OOXML Excel
        try:
            file.seek(0)
            return pd.read_excel(file)
        except Exception as e2:
            st.error(f"Napaka pri branju datoteke {file.name}: {e2}")
            return None
    return df


def extract_number(val):
    if pd.isna(val):
        return None
    m = re.search(r"\d+", str(val))
    return m.group() if m else None


def parse_date(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s:
        return None
    try:
        return pd.to_datetime(s).date()
    except Exception:
        return None


def process_file(file, filter_date, min_days, urgent_days) -> "pd.DataFrame | None":
    df = parse_file(file)
    if df is None:
        return None

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"Datoteka **{file.name}** nima pričakovanih stolpcev: {missing}")
        return None

    work = df[REQUIRED_COLS].copy()

    # razpar-anje datumov
    work["_datum_nastanka_parsed"] = work["Datum nastanka"].apply(parse_date)
    work["_prihod_parsed"] = work["Prihod"].apply(parse_date)
    work = work.dropna(subset=["_datum_nastanka_parsed"])

    # filter statusa - "Na čakanju" (case-insensitive, robustno na HTML/presledke)
    work["_status_clean"] = work["Status"].astype(str).str.strip()
    work = work[work["_status_clean"].str.contains("čakanju", case=False, na=False)]

    if work.empty:
        return work

    # dni od nastanka do datuma filtracije (dolgo čakanje)
    work["Dni od nastanka"] = work["_datum_nastanka_parsed"].apply(
        lambda d: (filter_date - d).days
    )
    dolgo_cakanje = work["Dni od nastanka"] >= min_days

    # dni med nastankom rezervacije in prihodom gosta (prihod kmalu = nujno,
    # ker gost mora plačati vnaprej, rok za urejanje je kratek)
    work["Dni do prihoda (od nastanka)"] = work.apply(
        lambda r: (r["_prihod_parsed"] - r["_datum_nastanka_parsed"]).days
        if pd.notna(r["_prihod_parsed"])
        else None,
        axis=1,
    )
    prihod_kmalu = work["Dni do prihoda (od nastanka)"].apply(
        lambda v: v is not None and v <= urgent_days
    )

    work = work[dolgo_cakanje | prihod_kmalu]
    if work.empty:
        return work

    def _razlog(row):
        r = []
        if row["Dni od nastanka"] >= min_days:
            r.append(f"Dolgo čakanje (≥{min_days} dni)")
        d = row["Dni do prihoda (od nastanka)"]
        if d is not None and d <= urgent_days:
            r.append(f"Prihod kmalu (≤{urgent_days} dni od nastanka)")
        return " + ".join(r)

    work["Razlog"] = work.apply(_razlog, axis=1)

    work["PMS koda"] = work["PMS koda"].apply(extract_number)
    work["Code"] = work["Code"].astype(str).str.strip()
    work["Objekt"] = work["Objekt"].astype(str).str.strip()
    work["Datum nastanka"] = work["_datum_nastanka_parsed"].astype(str)
    work["Prihod"] = work["Prihod"].astype(str).str.strip()
    work["Lastnik rezervacije"] = work["Lastnik rezervacije"].astype(str).str.strip()
    work["Vir datoteke"] = file.name

    final_cols = [
        "Code",
        "PMS koda",
        "Objekt",
        "Datum nastanka",
        "Prihod",
        "Lastnik rezervacije",
        "Status",
        "Dni od nastanka",
        "Dni do prihoda (od nastanka)",
        "Razlog",
        "Vir datoteke",
    ]
    return work[final_cols].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Glavna logika
# ---------------------------------------------------------------------------
if uploaded_files:
    if len(uploaded_files) > 6:
        st.warning("Naložiš lahko največ 6 datotek. Upoštevanih bo prvih 6.")
        uploaded_files = uploaded_files[:6]

    all_results = []
    for f in uploaded_files:
        res = process_file(f, filter_date, min_days, urgent_days)
        if res is not None and not res.empty:
            all_results.append(res)
            st.caption(f"✅ {f.name}: najdenih {len(res)} vrstic")
        elif res is not None:
            st.caption(f"⚪ {f.name}: ni vrstic, ki ustrezajo pogojem")

    if all_results:
        combined = pd.concat(all_results, ignore_index=True)
        st.success(f"Skupno najdenih {len(combined)} vrstic, ki ustrezajo pogojem.")

        urgent_mask = combined["Razlog"].astype(str).str.contains("Prihod kmalu")

        def _highlight_urgent(row):
            is_urgent = urgent_mask.loc[row.name]
            return ["background-color: #ffcccc" if is_urgent else "" for _ in row]

        st.caption("🔴 Rdeče označene vrstice = prihod je 1-3 dni (oz. nastavljeni prag) od nastanka rezervacije - nujno preveriti.")
        st.dataframe(
            combined.style.apply(_highlight_urgent, axis=1),
            use_container_width=True,
        )

        # Excel za prenos (z rdečo osvetlitvijo vrstic "prihod kmalu")
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            combined.to_excel(writer, index=False, sheet_name="Na čakanju")
            worksheet = writer.sheets["Na čakanju"]
            red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            n_cols = combined.shape[1]
            for excel_row, is_urgent in enumerate(urgent_mask, start=2):  # vrstica 1 = header
                if is_urgent:
                    for col in range(1, n_cols + 1):
                        worksheet.cell(row=excel_row, column=col).fill = red_fill
        output.seek(0)

        st.download_button(
            label="⬇️ Prenesi rezultate kot Excel (.xlsx)",
            data=output,
            file_name=f"na_cakanju_{filter_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Ni najdenih vrstic, ki bi ustrezale filtru v nobeni naloženi datoteki.")
else:
    st.info("Prosim, naloži vsaj eno XLS datoteko (do največ 6).")
