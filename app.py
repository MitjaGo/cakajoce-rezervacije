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
from urllib.parse import quote

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Rezervacije - Na čakanju", layout="wide")

LOGO_URL = "https://www.adria-ankaran.si//app/uploads/2025/10/logo-Adria.jpg"

header_left, header_right = st.columns([4, 1])
with header_left:
    st.title("📋 Filtriranje rezervacij iz PHOBSA s statusom NA ČAKANJU ")
with header_right:
    st.image(LOGO_URL, width=110)

st.markdowbsn(
    """pho
Naloži od **1 do 6** XLS datotek (izvoz iz sistema PHOBS / Rezervacije na čakanju ). 
*Označi na Phobsu preden prenseš v excel samo vrstice s statusom **Na čakanju** z rumeno ali oranžno obarvano podlago

Aplikacija bo:
- prebrala podatke iz vsake datoteke,
- izračunala, koliko dni je preteklo od stolpca **Datum nastanka** do izbranega
  datuma filtracije,
- prikazala vrstice, kjer je preteklo **N ali več dni** (privzeto 4),
- prikazala stolpce: **Številka PH, HIS, Objekt, Datum ponudbe, Prihod,
  Lastnik rezervacije, Status**,
- združila rezultate vseh naloženih datotek v en Excel dokument, ki ga
  prenesete na svoj računalnik.
"""
)

# ---------------------------------------------------------------------------
# Nastavitve filtra
# ---------------------------------------------------------------------------
URGENT_DAYS = 3  # rezervacije s prihodom 1, 2 ali 3 dni po nastanku - vedno prikazane

col1, col2 = st.columns(2)
with col1:
    filter_date = st.date_input("Datum filtracije", value=date.today())
with col2:
    min_days = st.number_input(
        "Min. dni od 'Datum nastanka' (dolgo čakanje)", min_value=0, value=4, step=1
    )

st.caption(
    "Vrstica se prikaže, če je status 'Na čakanju' IN (od nastanka je "
    f"preteklo ≥ zgornji prag DNI, ALI je prihod le {URGENT_DAYS} dni ali manj "
    "od nastanka rezervacije - gost mora plačati vnaprej, zato je treba te "
    "rezervacije nujno preveriti)."
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
    result = work[final_cols].reset_index(drop=True)

    # preimenovanje za prikaz/izvoz (interno branje iz XLS ostane nespremenjeno)
    result = result.rename(
        columns={
            "Code": "Številka PH",
            "PMS koda": "HIS",
            "Datum nastanka": "Datum ponudbe",
        }
    )
    return result


def _table_html_parts(df: pd.DataFrame, urgent_mask: pd.Series):
    """Vrne (header_html, rows_html) - skupna gradnja za tisk in kopiranje."""
    header_html = "".join(f"<th>{c}</th>" for c in df.columns)
    rows_html = ""
    for idx, row in df.iterrows():
        row_style = ' style="background-color:#ffcccc;"' if urgent_mask.loc[idx] else ""
        cells = "".join(f"<td>{'' if pd.isna(v) else v}</td>" for v in row)
        rows_html += f"<tr{row_style}>{cells}</tr>"
    return header_html, rows_html


def build_print_html(df: pd.DataFrame, urgent_mask: pd.Series, filter_date) -> str:
    """Zgradi samostojen HTML dokument s tabelo, oblikovan za tiskanje na A4,
    z gumbom, ki sproži tiskanje (window.print())."""
    header_html, rows_html = _table_html_parts(df, urgent_mask)

    return f"""
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        @page {{ size: A4 landscape; margin: 12mm; }}
        body {{ font-family: Arial, Helvetica, sans-serif; }}
        h2 {{ margin: 0 0 4px 0; }}
        p.meta {{ margin: 0 0 12px 0; color: #555; font-size: 12px; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #999; padding: 4px 6px; font-size: 10px; text-align: left; }}
        th {{ background-color: #f0f0f0; }}
        #printBtn {{
            padding: 8px 18px; font-size: 14px; cursor: pointer;
            background-color: #d63333; color: white; border: none; border-radius: 4px;
        }}
        @media print {{
            #printBtn {{ display: none; }}
        }}
    </style>
    </head>
    <body>
        <button id="printBtn" onclick="window.print()">🖨️ Natisni na A4</button>
        <h2>Rezervacije - Na čakanju</h2>
        <p class="meta">Datum filtracije: {filter_date} · Skupno vrstic: {len(df)}</p>
        <table>
            <thead><tr>{header_html}</tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
    </body>
    </html>
    """


def build_table_html_for_clipboard(df: pd.DataFrame, urgent_mask: pd.Series, filter_date) -> str:
    """Zgradi HTML tabelo (brez gumbov/strani), primerno za kopiranje v
    odložišče in lepljenje neposredno v telo e-maila (npr. Outlook), kjer se
    prikaže kot prava, oblikovana tabela - enako kot pri tisku."""
    header_html, rows_html = _table_html_parts(df, urgent_mask)
    return (
        f'<div style="font-family:Arial,Helvetica,sans-serif;">'
        f'<h3 style="margin:0 0 4px 0;">Rezervacije - Na čakanju</h3>'
        f'<p style="margin:0 0 10px 0;color:#555;font-size:12px;">'
        f'Datum filtracije: {filter_date} &middot; Skupno vrstic: {len(df)}</p>'
        f'<table style="border-collapse:collapse;width:100%;">'
        f'<thead><tr>{header_html.replace("<th>", "<th style=\'border:1px solid #999;padding:4px 6px;background:#f0f0f0;font-size:11px;text-align:left;\'>")}</tr></thead>'
        f'<tbody>{rows_html.replace("<td>", "<td style=\'border:1px solid #999;padding:4px 6px;font-size:11px;\'>")}</tbody>'
        f'</table></div>'
    )


def build_table_text_for_clipboard(df: pd.DataFrame) -> str:
    """Navadno-tekstovna (tab-ločena) različica tabele - kot rezervni format
    za odložišče (npr. za lepljenje v Excel)."""
    lines = ["\t".join(str(c) for c in df.columns)]
    for _, row in df.iterrows():
        lines.append("\t".join("" if pd.isna(v) else str(v) for v in row))
    return "\n".join(lines)


def build_mailto_link(df: pd.DataFrame, filter_date, recipient: str = "") -> str:
    """Zgradi mailto: povezavo. Ker mailto ne podpira HTML telesa, doda
    napotek za lepljenje predhodno kopirane tabele (gumb 'Kopiraj tabelo'),
    pod njim pa še preprost tekstovni povzetek kot rezervo."""
    subject = f"Rezervacije na čakanju - {filter_date}"
    lines = [
        "Tukaj prilepi tabelo (Ctrl+V) - najprej klikni gumb 'Kopiraj tabelo':",
        "",
        "",
        "---",
        f"Rezervacije s statusom 'Na čakanju' na dan {filter_date} "
        f"(skupno {len(df)} vrstic) - tekstovni povzetek:",
        "",
    ]
    for _, row in df.iterrows():
        lines.append(
            f"{row['Številka PH']} | {row['HIS']} | {row['Objekt']} | "
            f"Ponudba: {row['Datum ponudbe']} | Prihod: {row['Prihod']} | "
            f"{row['Lastnik rezervacije']} | {row['Razlog']}"
        )
    body = "\n".join(lines)
    return f"mailto:{recipient}?subject={quote(subject)}&body={quote(body)}"


# ---------------------------------------------------------------------------
# Glavna logika
# ---------------------------------------------------------------------------
if uploaded_files:
    if len(uploaded_files) > 6:
        st.warning("Naložiš lahko največ 6 datotek. Upoštevanih bo prvih 6.")
        uploaded_files = uploaded_files[:6]

    all_results = []
    for f in uploaded_files:
        res = process_file(f, filter_date, min_days, URGENT_DAYS)
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

        btn_col1, btn_col2, btn_col3, btn_col4 = st.columns(4)
        with btn_col1:
            st.download_button(
                label="⬇️ Prenesi kot Excel (.xlsx)",
                data=output,
                file_name=f"na_cakanju_{filter_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with btn_col2:
            print_html = build_print_html(combined, urgent_mask, filter_date)
            components.html(
                f"""
                <div style="display:flex; justify-content:center;">
                    <button onclick="printTable()" style="
                        width:100%; padding:8px 0; font-size:14px; cursor:pointer;
                        background-color:#31333F; color:white; border:none;
                        border-radius:6px;">🖨️ Natisni (A4)</button>
                </div>
                <script>
                function printTable() {{
                    var w = window.open('', '_blank');
                    w.document.write({print_html!r});
                    w.document.close();
                    w.focus();
                    setTimeout(function() {{ w.print(); }}, 300);
                }}
                </script>
                """,
                height=45,
            )
        with btn_col3:
            copy_html = build_table_html_for_clipboard(combined, urgent_mask, filter_date)
            copy_text = build_table_text_for_clipboard(combined)
            components.html(
                f"""
                <div style="display:flex; justify-content:center;">
                    <button id="copyBtn" onclick="copyTable()" style="
                        width:100%; padding:8px 0; font-size:14px; cursor:pointer;
                        background-color:#31333F; color:white; border:none;
                        border-radius:6px;">📋 Kopiraj tabelo</button>
                </div>
                <script>
                async function copyTable() {{
                    var btn = document.getElementById('copyBtn');
                    var htmlContent = {copy_html!r};
                    var textContent = {copy_text!r};
                    try {{
                        var item = new ClipboardItem({{
                            'text/html': new Blob([htmlContent], {{type: 'text/html'}}),
                            'text/plain': new Blob([textContent], {{type: 'text/plain'}})
                        }});
                        await navigator.clipboard.write([item]);
                        btn.innerText = '✅ Kopirano!';
                    }} catch (err) {{
                        btn.innerText = '⚠️ Kopiranje ni uspelo';
                        console.error(err);
                    }}
                    setTimeout(function() {{ btn.innerText = '📋 Kopiraj tabelo'; }}, 2500);
                }}
                </script>
                """,
                height=45,
            )
        with btn_col4:
            mailto_url = build_mailto_link(
                combined, filter_date, recipient="mitja.goja@adria-ankaran.si"
            )
            st.link_button(
                "📧 Pošlji kot e-mail",
                url=mailto_url,
                use_container_width=True,
            )
        st.caption(
            "💡 Za lepo oblikovano tabelo v e-mailu: najprej klikni **'📋 Kopiraj "
            "tabelo'**, nato **'📧 Pošlji kot e-mail'** (odpre Outlook) in v telo "
            "e-maila prilepi (Ctrl+V) - tabela se prilepi enako oblikovana kot pri "
            "tisku. Excel priloge zaradi omejitev brskalnika ni mogoče samodejno "
            "pripeti - za to najprej prenesi Excel in ga ročno priloži e-mailu."

        )
    else:
        st.info("Ni najdenih vrstic, ki bi ustrezale filtru v nobeni naloženi datoteki.")
else:
    st.info("Prosim, naloži vsaj eno XLS datoteko (do največ 6).")






