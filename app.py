import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
from fpdf import FPDF
import os

st.set_page_config(page_title="MR Fakturagenerator (Ajour + Dansk Omsorgspleje + Dit Vikarbureau)", layout="centered")

# ---------- Styling ----------
st.markdown(
    """
<style>
body { background-color:#aa1e1e; color:white; }
[data-testid="stAppViewContainer"] > .main {
    background:white;
    color:black;
    border-radius:10px;
    padding:2rem;
    max-width:980px;
    margin:auto;
}
</style>
""",
    unsafe_allow_html=True,
)

if os.path.exists("logo.png"):
    st.image("logo.png", width=80)

# --------------------------------------------------
# CONSTANTS (FROM / TO)
# --------------------------------------------------
FROM_INFO = {
    "name": "MR Rekruttering",
    "addr": "Valbygårdsvej 1, 4. th, 2500 Valby",
    "cvr": "45090965",
    "phone": "71747290",
    "web": "www.akutvikar.com",
}

BANK_INFO_LINE1 = "Bank: Finseta | IBAN: GB79TCCL04140404627601 | BIC: TCCLGB3LXXX"
BANK_INFO_LINE2 = "Betalingsbetingelser: Bankoverførsel. Fakturanr. bedes angivet ved betaling."

TO_AJOUR = {
    "title": "Ajour Care ApS",
    "lines": [
        "CVR: 34478953",
        "Kontakt: Charlotte Bigum Christensen",
        "Email: cbc@ajourcare.dk",
    ],
}

TO_DANSK = {
    "title": "DANSK OMSORGSPLEJE APS",
    "lines": [
        "CVR: 42092630",
        "Frederiksborgvej 14, st, 3200 Helsinge",
    ],
}

TO_DIT = {
    "title": "Dit Vikarbureau",
    "lines": [
        "CVR: (indsæt CVR)",
        "Adresse: (indsæt adresse)",
        "Email: (indsæt email)",
    ],
}

# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def normalize_personale(val: str) -> str:
    if val is None:
        return ""
    s = str(val).replace("\u00A0", " ").strip().lower()
    s = " ".join(s.split())
    if s == "assistent 2":
        s = "assistent"
    if "ufagl" in s:
        return "ufaglært"
    if "hjælp" in s:
        return "hjælper"
    if "assist" in s:
        return "assistent"
    if "sygepl" in s:
        return "sygeplejerske"
    return s


def ensure_datetime(series) -> pd.Series:
    return pd.to_datetime(series, dayfirst=True, errors="coerce")


def safe_time_str(val) -> str:
    if pd.isna(val):
        return ""
    s = str(val)
    if len(s) >= 5 and s[2] == ":":
        return s[:5]
    return s[:5]


def build_tidsperiode(start, end) -> str:
    return f"{safe_time_str(start)}-{safe_time_str(end)}"


def parse_start_time_to_minutes(tidsperiode: str) -> int:
    """
    Parse start time from 'HH:MM-HH:MM' into minutes since 00:00.
    Returns 0 if missing/bad.
    """
    try:
        s = str(tidsperiode).split("-")[0].strip()
        hh, mm = s.split(":")
        return int(hh) * 60 + int(mm)
    except Exception:
        return 0


def time_to_hour(t: str) -> int:
    try:
        return int(str(t)[:2])
    except Exception:
        return 0


# --------------------------------------------------
# BASE CLEANING (used for all)
# --------------------------------------------------
def rens_data_base(df: pd.DataFrame) -> pd.DataFrame:
    # ✅ KEEP DitVikar rows so you can invoice them too

    needed = [
        "Dato",
        "Medarbejder",
        "Starttid",
        "Sluttid",
        "Timer",
        "Personalegruppe",
        "Jobfunktion",
        "Shift status",
        "Afdeling",
    ]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Mangler kolonner i filen: {', '.join(missing)}")

    df = df[needed].copy()

    # Remove 0/blank hours
    df["Timer"] = pd.to_numeric(df["Timer"], errors="coerce")
    df = df[df["Timer"].notna() & (df["Timer"] > 0)].copy()

    # Parse date
    df["Dato"] = ensure_datetime(df["Dato"])
    df = df[df["Dato"].notna()].copy()

    # Tidsperiode + raw jobfunction
    df["Tidsperiode"] = df.apply(lambda r: build_tidsperiode(r["Starttid"], r["Sluttid"]), axis=1)
    df["Jobfunktion_raw"] = df["Jobfunktion"]

    # Normalize personale
    df["Personale"] = df["Personalegruppe"].apply(normalize_personale)

    # For proper chronological sorting inside Jobfunktion
    df["StartMin"] = df["Tidsperiode"].apply(parse_start_time_to_minutes)

    return df


# --------------------------------------------------
# AJOURCARE: jobfunktion mapping (your logic)
# --------------------------------------------------
def map_jobfunktion_ajour(df: pd.DataFrame) -> pd.DataFrame:
    byer = ["allerød", "egedal", "frederiksund", "solrød", "herlev", "ringsted", "køge"]

    def find_by(txt):
        t = str(txt).lower()
        for b in byer:
            if b in t:
                return b
        if "kirsten" in t:
            return "køge"
        return "andet"

    out = df.copy()
    out["Jobfunktion"] = out["Jobfunktion"].apply(find_by)
    return out


# --------------------------------------------------
# DANSK OMSORGSPLEJE: jobfunktion display (simple extract)
# --------------------------------------------------
def extract_location_dansk(jobfunction):
    if not jobfunction:
        return ""
    parts = str(jobfunction).split("-")
    return parts[1].strip() if len(parts) > 1 else str(jobfunction).strip()


def extract_location_dit(jobfunction):
    # same display rule as Dansk (change if needed)
    return extract_location_dansk(jobfunction)


# --------------------------------------------------
# RATE LOGIC
# --------------------------------------------------
# ✅ KEEP your existing Ajour/AkutVikar logic unchanged
def beregn_takst_ajour(row) -> int:
    helligdag = row["Helligdag"] == "Ja"
    personale = row["Personale"]

    start_hour = time_to_hour(row["Tidsperiode"].split("-")[0])
    dag = start_hour < 15
    weekend = row["Dato"].weekday() >= 5

    if personale == "ufaglært":
        if helligdag: return 215 if dag else 220
        return 215 if weekend and dag else 220 if weekend else 175 if dag else 210

    if personale == "hjælper":
        if helligdag: return 215 if dag else 220
        return 215 if weekend and dag else 220 if weekend else 200 if dag else 210

    if personale == "assistent":
        if helligdag: return 230 if dag else 240
        return 230 if weekend and dag else 240 if weekend else 220 if dag else 225

    return 0


# ✅ KEEP your existing Dansk logic unchanged
def beregn_takst_dansk(row) -> int:
    if row["Helligdag"] == "Ja":
        return 350
    weekend = row["Dato"].weekday() >= 5
    if weekend:
        return 300
    start_hour = time_to_hour(row["Tidsperiode"].split("-")[0])
    return 280 if start_hour >= 15 else 255


# ✅ NEW: DitVikar udbudspriser from your screenshot (excl. moms)
DITVIKAR_RATES = {
    "hjælper": {
        "weekday_day": 333.00,
        "weekday_night": 406.26,
        "weekend_day": 499.50,
        "weekend_night": 572.76,
        "holiday_day": 666.00,
        "holiday_night": 739.26,
    },
    "assistent": {
        "weekday_day": 353.00,
        "weekday_night": 430.66,
        "weekend_day": 529.50,
        "weekend_night": 607.16,
        "holiday_day": 706.00,
        "holiday_night": 783.66,
    },
    "sygeplejerske": {
        "weekday_day": 386.00,
        "weekday_night": 482.50,
        "weekend_day": 579.00,
        "weekend_night": 675.50,
        "holiday_day": 772.00,
        "holiday_night": 868.50,
    },
}

def is_day_window(start_min: int) -> bool:
    """
    Day window in your price sheet: 06:00–15:00
    Everything else is evening/night: 15:00–05:00
    """
    return 6 * 60 <= start_min < 15 * 60

def beregn_takst_dit(row) -> float:
    personale = row["Personale"]
    rates = DITVIKAR_RATES.get(personale)
    if not rates:
        # If you later add ufaglært prices, put them in DITVIKAR_RATES
        return 0.0

    helligdag = row["Helligdag"] == "Ja"
    weekend = row["Dato"].weekday() >= 5
    start_min = parse_start_time_to_minutes(row["Tidsperiode"])
    day = is_day_window(start_min)

    if helligdag:
        return rates["holiday_day"] if day else rates["holiday_night"]
    if weekend:
        return rates["weekend_day"] if day else rates["weekend_night"]
    return rates["weekday_day"] if day else rates["weekday_night"]


# --------------------------------------------------
# PDF GENERATION (generic)
# --------------------------------------------------
def generer_pdf(inv: pd.DataFrame, fakturanr: int, to_info: dict, filename_prefix: str) -> tuple[BytesIO, str]:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=18)

    if os.path.exists("logo.png"):
        pdf.image("logo.png", 10, 5, 30)

    pdf.set_font("Arial", "B", 20)
    pdf.set_xy(140, 10)
    pdf.cell(60, 10, f"FAKTURA {fakturanr}", align="R")

    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(10, 40)
    pdf.cell(95, 6, f"Fra: {FROM_INFO['name']}", ln=1)
    pdf.set_font("Arial", "", 10)
    pdf.set_x(10); pdf.cell(95, 6, FROM_INFO["addr"], ln=1)
    pdf.set_x(10); pdf.cell(95, 6, f"CVR.nr. {FROM_INFO['cvr']}", ln=1)
    pdf.set_x(10); pdf.cell(95, 6, f"Tlf: {FROM_INFO['phone']}", ln=1)
    pdf.set_x(10); pdf.cell(95, 6, f"Web: {FROM_INFO['web']}", ln=1)

    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(105, 40)
    pdf.cell(95, 6, f"Til: {to_info['title']}", ln=1)
    pdf.set_font("Arial", "", 10)
    y = pdf.get_y()
    for line in to_info["lines"]:
        pdf.set_xy(105, y)
        pdf.cell(95, 6, line, ln=1)
        y = pdf.get_y()

    pdf.ln(4)
    pdf.set_x(10)
    pdf.cell(0, 6, f"Fakturadato: {date.today().strftime('%d.%m.%Y')}", ln=1)
    pdf.ln(4)

    cols = ["Dato","Medarbejder","Tidsperiode","Timer","Personale","Jobfunktion","Helligdag","Takst","Samlet"]
    widths = [18, 32, 22, 10, 16, 30, 14, 14, 14]  # allow decimals in Takst

    pdf.set_font("Arial","B",8)
    pdf.set_x(10)
    for h,w in zip(cols, widths):
        pdf.cell(w,8,h,1,align='C')
    pdf.ln()

    pdf.set_font("Arial","",8)
    total = 0.0
    for _, r in inv.iterrows():
        pdf.set_x(10)
        row = [
            r["Dato"].strftime("%d.%m.%Y"),
            str(r["Medarbejder"])[:26],
            r["Tidsperiode"],
            f"{float(r['Timer']):.1f}",
            str(r["Personale"])[:12],
            str(r["Jobfunktion"])[:14],
            r["Helligdag"],
            f"{float(r['Takst']):.2f}",     # ✅ decimals
            f"{float(r['Samlet']):.2f}",
        ]
        for v,w in zip(row, widths):
            pdf.cell(w,8,str(v),1,align='C')
        pdf.ln()
        total += float(r["Samlet"])

    moms = total * 0.25
    pdf.ln(4)
    pdf.set_font("Arial","B",10)
    pdf.cell(0,6,f"Subtotal: {total:.2f} kr",ln=1)
    pdf.cell(0,6,f"Moms (25%): {moms:.2f} kr",ln=1)
    pdf.cell(0,6,f"Total inkl. moms: {total+moms:.2f} kr",ln=1)

    pdf.ln(5)
    pdf.set_font("Arial","",9)
    pdf.cell(0,6,BANK_INFO_LINE1,ln=1)
    pdf.cell(0,6,BANK_INFO_LINE2,ln=1)

    pdf_bytes = pdf.output(dest="S").encode("latin-1")
    out = BytesIO(pdf_bytes)
    out.seek(0)
    return out, f"{filename_prefix}_{fakturanr}.pdf"


# --------------------------------------------------
# BUILD INVOICE DF (per customer) + SORTING
# --------------------------------------------------
def build_invoice_df(df_customer: pd.DataFrame, helligdage: list[pd.Timestamp], rate_func) -> pd.DataFrame:
    inv = df_customer.copy()

    hellig_set = set(pd.to_datetime(helligdage))
    inv["Helligdag"] = inv["Dato"].dt.normalize().isin([h.normalize() for h in hellig_set]).map({True:"Ja", False:"Nej"})

    inv["Takst"] = [rate_func(r) for _, r in inv.iterrows()]
    inv["Samlet"] = inv["Timer"] * inv["Takst"]

    inv = inv[["Dato","Medarbejder","Tidsperiode","StartMin","Timer","Personale","Jobfunktion","Helligdag","Takst","Samlet"]].copy()
    inv = inv.sort_values(["Jobfunktion","Dato","StartMin","Medarbejder"], ascending=[True,True,True,True])
    inv = inv.drop(columns=["StartMin"])
    return inv


# --------------------------------------------------
# UI
# --------------------------------------------------
st.title("MR Rekruttering – Fakturagenerator (Ajour Care + Dansk Omsorgspleje + Dit Vikarbureau)")

file = st.file_uploader("Upload vagtplan-fil (Excel)", type=["xlsx","xls"])

c1, c2, c3 = st.columns(3)
with c1:
    faktura_ajour = st.number_input("Fakturanummer (Ajour Care / AkutVikar)", min_value=0, step=1, value=0)
with c2:
    faktura_dansk = st.number_input("Fakturanummer (Dansk Omsorgspleje)", min_value=0, step=1, value=0)
with c3:
    faktura_dit = st.number_input("Fakturanummer (Dit Vikarbureau)", min_value=0, step=1, value=0)

if file:
    try:
        raw = pd.read_excel(file)
        clean = rens_data_base(raw)

        afdeling_lower = clean["Afdeling"].astype(str).str.lower()

        df_ajour = clean[
            afdeling_lower.str.contains(r"ajour|akut\s*-?\s*vikar|akutvikar", regex=True, na=False)
        ].copy()

        df_dansk = clean[
            afdeling_lower.str.contains(r"dansk\s*omsorgspleje", regex=True, na=False)
        ].copy()

        # Dit Vikarbureau (supports misspelling "buerou" too)
        df_dit = clean[
            afdeling_lower.str.contains(r"dit\s*vikar|ditvikar|dit\s*vikarbureau|dit\s*vikarbuerou", regex=True, na=False)
        ].copy()

        if len(df_ajour) > 0:
            df_ajour = map_jobfunktion_ajour(df_ajour)

        if len(df_dansk) > 0:
            df_dansk["Jobfunktion"] = df_dansk["Jobfunktion"].apply(extract_location_dansk)

        if len(df_dit) > 0:
            df_dit["Jobfunktion"] = df_dit["Jobfunktion"].apply(extract_location_dit)

        st.markdown("### Fil-status")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Rækker (total)", len(clean))
        m2.metric("Ajour/Akut Vikar rækker", len(df_ajour))
        m3.metric("Dansk Omsorgspleje rækker", len(df_dansk))
        m4.metric("Dit Vikarbureau rækker", len(df_dit))

        all_dates = sorted(clean["Dato"].dt.date.unique())
        helligdage = [pd.Timestamp(d) for d in st.multiselect("Vælg helligdage (gælder for alle)", all_dates)]

        can_generate = True
        if len(df_ajour) > 0 and faktura_ajour <= 0:
            can_generate = False
            st.warning("Du har Ajour/Akut Vikar-rækker, men mangler fakturanummer til Ajour/AkutVikar.")
        if len(df_dansk) > 0 and faktura_dansk <= 0:
            can_generate = False
            st.warning("Du har Dansk Omsorgspleje-rækker, men mangler fakturanummer til Dansk Omsorgspleje.")
        if len(df_dit) > 0 and faktura_dit <= 0:
            can_generate = False
            st.warning('Du har Dit Vikarbureau-rækker, men mangler fakturanummer til "Dit Vikarbureau".')

        if st.button("Generer PDF'er", disabled=not can_generate):
            st.success("PDF'er genereres…")

            # AJOUR / AKUTVIKAR (unchanged rates)
            if len(df_ajour) > 0:
                inv_ajour = build_invoice_df(df_ajour, helligdage, beregn_takst_ajour)

                # Kirsten +10 (based on Jobfunktion_raw in df_ajour)
                kirsten_flag = df_ajour[["Dato", "Medarbejder", "Tidsperiode", "Jobfunktion_raw"]].copy()
                kirsten_flag["Kirsten"] = kirsten_flag["Jobfunktion_raw"].astype(str).str.contains("kirsten", case=False, na=False)
                kirsten_flag = kirsten_flag.drop(columns=["Jobfunktion_raw"])

                inv_ajour = inv_ajour.merge(kirsten_flag, on=["Dato", "Medarbejder", "Tidsperiode"], how="left")
                inv_ajour["Kirsten"] = inv_ajour["Kirsten"].fillna(False)

                inv_ajour.loc[inv_ajour["Kirsten"] == True, "Takst"] = inv_ajour["Takst"] + 10
                inv_ajour["Samlet"] = inv_ajour["Timer"] * inv_ajour["Takst"]
                inv_ajour = inv_ajour.drop(columns=["Kirsten"])

                pdf_ajour, pdf_ajour_name = generer_pdf(inv_ajour, int(faktura_ajour), TO_AJOUR, "Faktura_AjourCare")
                st.download_button("Download PDF (Ajour/AkutVikar)", pdf_ajour, file_name=pdf_ajour_name)

            # DANSK (unchanged rates)
            if len(df_dansk) > 0:
                inv_dansk = build_invoice_df(df_dansk, helligdage, beregn_takst_dansk)
                pdf_dansk, pdf_dansk_name = generer_pdf(inv_dansk, int(faktura_dansk), TO_DANSK, "Faktura_DanskOmsorgspleje")
                st.download_button("Download PDF (Dansk Omsorgspleje)", pdf_dansk, file_name=pdf_dansk_name)

            # DIT VIKARBUREAU (new rates from image)
            if len(df_dit) > 0:
                inv_dit = build_invoice_df(df_dit, helligdage, beregn_takst_dit)
                pdf_dit, pdf_dit_name = generer_pdf(inv_dit, int(faktura_dit), TO_DIT, "Faktura_DitVikarbureau")
                st.download_button("Download PDF (Dit Vikarbureau)", pdf_dit, file_name=pdf_dit_name)

            if len(df_ajour) == 0 and len(df_dansk) == 0 and len(df_dit) == 0:
                st.error('Ingen rækker fundet for Ajour/Akut Vikar, Dansk Omsorgspleje eller Dit Vikarbureau i Afdeling-kolonnen.')

    except Exception as e:
        st.error(f"Fejl: {e}")

