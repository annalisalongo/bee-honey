import io
import json
import uuid
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Bee Honey", page_icon="🐝", layout="wide")

DATA_DIR = Path("data")
IMG_DIR = DATA_DIR / "images"
DATA_DIR.mkdir(exist_ok=True)
IMG_DIR.mkdir(exist_ok=True)

FILES = {
    "controlli": DATA_DIR / "controlli.csv",
    "acquisti": DATA_DIR / "acquisti_magazzino.csv",
    "invasettato": DATA_DIR / "miele_invasettato.csv",
    "vendite": DATA_DIR / "vendite_miele.csv",
    "watchlist": DATA_DIR / "bee_deals_watchlist.csv",
    "offerte": DATA_DIR / "bee_deals_offerte.csv",
}

ARNIE = ["Bee Bianca", "Bee Verdina", "Bee Gialla", "Bee Verde"]
FORZE = ["Debole", "Media", "Forte", "Fortissima"]
NUTRIZIONI = ["Nessuna", "Candito", "Sciroppo", "Altro"]
MAG_CATEGORIE = ["Fogli cerei", "Telaini", "Melari", "Barattoli", "Nutrizione", "Attrezzatura", "Altro"]
UNITA = ["pz", "kg", "confezioni", "litri", "altro"]

st.markdown("""
<style>
.main .block-container {padding-top:1rem;padding-bottom:2rem;max-width:1180px;}
.hero {background:linear-gradient(135deg,#fff8e8 0%,#fffdf7 100%);border:1px solid #eadfbe;border-radius:22px;padding:20px;margin-bottom:18px;}
.card {border:1px solid #eadfbe;background:#fffdf8;border-radius:18px;padding:16px;margin-bottom:12px;}
.kpi {display:inline-block;padding:4px 10px;border-radius:999px;border:1px solid #eadfbe;background:#fff;margin-right:6px;margin-bottom:6px;font-size:0.9rem;}
.bee-ok {background:#e8f7ee;border-left:6px solid #2c9b62;padding:12px;border-radius:12px;margin-bottom:10px;}
.bee-warn {background:#fff4df;border-left:6px solid #c98b00;padding:12px;border-radius:12px;margin-bottom:10px;}
.bee-danger {background:#fdecec;border-left:6px solid #c94b4b;padding:12px;border-radius:12px;margin-bottom:10px;}
.small {color:#6b604f;font-size:0.92rem;}
.section-title {font-size:1.08rem;font-weight:700;margin:10px 0 12px 0;}
div[data-testid="stMetric"] {background:#fff;border:1px solid #eadfbe;border-radius:18px;padding:8px 12px;}
div.stButton > button, .stDownloadButton button {width:100%;border-radius:14px;border:1px solid #d9b65d;background:#fff8e6;min-height:46px;font-weight:600;}
</style>
""", unsafe_allow_html=True)

CONTROLLI_COLS = ["id","arnia","data_controllo","forza_colonia","telaini_coperti","covata_fresca","covata_opercolata","celle_reali","regina_vista","regina_nuova","melario_presente","melario_percento","nutrizione","api_nervose","note","prossimo_controllo","foto"]
ACQUISTI_COLS = ["id","data_acquisto","categoria","prodotto","quantita","unita_misura","prezzo_totale","fornitore_sito","note"]
INVASETTATO_COLS = ["id","data_invasettamento","peso_grammi","quantita_invasettata","lotto","note"]
VENDITE_COLS = ["id","data_vendita","peso_grammi","quantita_venduta","prezzo_unitario","canale_vendita","note"]
WATCHLIST_COLS = ["id","prodotto","categoria","prezzo_target","negozio_preferito","note"]
OFFERTE_COLS = ["id","data_offerta","prodotto","sito","url","prezzo","spedizione","prezzo_totale","prezzo_target","valutazione","note"]

def load_csv(path: Path, cols):
    if path.exists():
        df = pd.read_csv(path)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        return df[cols]
    return pd.DataFrame(columns=cols)

def save_csv(df, path: Path):
    df.to_csv(path, index=False)

def as_bool(v):
    return str(v) == "True"

def to_num(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)

def euro(value):
    try:
        return f"€ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "€ 0,00"

def save_uploaded_files(files, folder_name: str, data_rif: str):
    paths = []
    if not files:
        return ""
    folder = IMG_DIR / folder_name.lower().replace(" ", "_") / data_rif
    folder.mkdir(parents=True, exist_ok=True)
    for f in files:
        ext = Path(f.name).suffix.lower() or ".jpg"
        out = folder / f"{uuid.uuid4().hex}{ext}"
        with open(out, "wb") as fp:
            fp.write(f.getbuffer())
        paths.append(str(out))
    return json.dumps(paths, ensure_ascii=False)

def analyze_control(row):
    alerts, suggestions, next_days = [], [], []
    forza = str(row.get("forza_colonia", ""))
    telaini = pd.to_numeric(pd.Series([row.get("telaini_coperti", 0)]), errors="coerce").fillna(0).iloc[0]
    covata_fresca = as_bool(row.get("covata_fresca", ""))
    celle_reali = as_bool(row.get("celle_reali", ""))
    regina_vista = as_bool(row.get("regina_vista", ""))
    regina_nuova = as_bool(row.get("regina_nuova", ""))
    melario = as_bool(row.get("melario_presente", ""))
    melario_pct = pd.to_numeric(pd.Series([row.get("melario_percento", 0)]), errors="coerce").fillna(0).iloc[0]
    api_nervose = as_bool(row.get("api_nervose", ""))
    if celle_reali and (forza in ["Forte", "Fortissima"] or telaini >= 8):
        alerts.append(("danger", "Rischio sciamatura alto: celle reali + colonia forte."))
        suggestions.append("Valuta divisione o controllo stretto delle celle reali.")
        next_days.append(3)
    if regina_nuova and not regina_vista:
        alerts.append(("ok", "Regina nuova o sospetta nuova: evita controlli troppo ravvicinati."))
        suggestions.append("Lascia tranquilla la colonia e fai un controllo mirato tra 7 e 10 giorni.")
        next_days.append(8)
    if melario and melario_pct >= 70:
        alerts.append(("warn", "Melario avanzato: prepara altro spazio sopra."))
        suggestions.append("Aggiungi spazio prima che il melario sia pieno del tutto.")
        next_days.append(4)
    if api_nervose:
        alerts.append(("warn", "Api nervose segnalate."))
        suggestions.append("Al prossimo controllo verifica meteo e manipolazione più rapida.")
        next_days.append(6)
    if not covata_fresca and not regina_vista and not regina_nuova:
        alerts.append(("warn", "Possibile problema regina: niente covata fresca e regina non vista."))
        suggestions.append("Controlla presenza di uova fresche al prossimo giro.")
        next_days.append(5)
    if not alerts:
        alerts.append(("ok", "Controllo nella norma."))
        suggestions.append("Continua con monitoraggio regolare.")
        next_days.append(7)
    return alerts, suggestions, str(date.today() + timedelta(days=min(next_days)))

def latest_controls(df):
    if df.empty:
        return pd.DataFrame(columns=CONTROLLI_COLS)
    tmp = df.copy()
    tmp["data_controllo_dt"] = pd.to_datetime(tmp["data_controllo"], errors="coerce")
    tmp = tmp.sort_values("data_controllo_dt", ascending=False)
    out = tmp.groupby("arnia", as_index=False).first()
    return out

def open_hive(hive_name: str):
    st.session_state["selected_hive"] = hive_name
    st.session_state["page"] = "Scheda arnia"
    st.rerun()

controlli = load_csv(FILES["controlli"], CONTROLLI_COLS)
acquisti = load_csv(FILES["acquisti"], ACQUISTI_COLS)
invasettato = load_csv(FILES["invasettato"], INVASETTATO_COLS)
vendite = load_csv(FILES["vendite"], VENDITE_COLS)
watchlist = load_csv(FILES["watchlist"], WATCHLIST_COLS)
offerte = load_csv(FILES["offerte"], OFFERTE_COLS)

if "selected_hive" not in st.session_state:
    st.session_state["selected_hive"] = ARNIE[0]
if "page" not in st.session_state:
    st.session_state["page"] = "Home"

pages = ["Home", "Overview arnie", "Scheda arnia", "Overview magazzino", "Nuovo controllo", "Consiglio AI", "Magazzino", "Export Excel"]
page = st.sidebar.selectbox("Menu", pages, key="page")

st.title("🐝 Bee Honey")
st.caption("Apiario, overview arnie, overview magazzino e scheda dedicata per ogni arnia.")

if page == "Home":
    latest = latest_controls(controlli)
    total_spese = float(to_num(acquisti.get("prezzo_totale", pd.Series(dtype=float))).sum()) if not acquisti.empty else 0.0
    st.markdown(f"""
    <div class="hero">
      <h2 style="margin:0;">Bee Honey Dashboard</h2>
      <div class="small" style="margin-top:6px;">Una panoramica veloce su apiario e magazzino.</div>
      <div style="margin-top:12px;">
        <span class="kpi">Arnie: {len(ARNIE)}</span>
        <span class="kpi">Controlli: {len(controlli)}</span>
        <span class="kpi">Articoli magazzino: {len(acquisti)}</span>
        <span class="kpi">Spese: {euro(total_spese)}</span>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('<div class="section-title">Le tue arnie</div>', unsafe_allow_html=True)
    cols = st.columns(2)
    for i, hive in enumerate(ARNIE):
        with cols[i % 2]:
            row = latest[latest["arnia"] == hive]
            if row.empty:
                st.markdown(f'<div class="card"><div style="font-weight:700;">{hive}</div><div class="small">Nessun dato</div></div>', unsafe_allow_html=True)
            else:
                r = row.iloc[0]
                notes = "" if str(r.get("note", "")) == "nan" else str(r.get("note", ""))
                notes = notes[:100] + ("..." if len(notes) > 100 else "")
                st.markdown(
                    f"""
                    <div class="card">
                      <div style="font-weight:700;font-size:1.04rem;">{hive}</div>
                      <div class="kpi">Ultimo: {r['data_controllo']}</div>
                      <div class="kpi">Forza: {r['forza_colonia']}</div>
                      <div class="kpi">Telaini: {r['telaini_coperti']}</div>
                      <div class="small">{notes}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            if st.button(f"Apri scheda {hive}", key=f"home_{hive}"):
                open_hive(hive)

if page == "Overview arnie":
    st.subheader("Overview arnie")
    latest = latest_controls(controlli)
    if latest.empty:
        st.info("Nessun controllo disponibile.")
    else:
        latest["telaini_num"] = to_num(latest["telaini_coperti"])
        latest["melario_num"] = to_num(latest["melario_percento"])
        st.dataframe(latest[["arnia","data_controllo","forza_colonia","telaini_coperti","celle_reali","regina_nuova","melario_presente","melario_percento","prossimo_controllo"]], use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### Telaini coperti")
            st.bar_chart(latest.set_index("arnia")[["telaini_num"]])
        with c2:
            st.markdown("### Melario %")
            st.bar_chart(latest.set_index("arnia")[["melario_num"]])
        st.markdown("### Apri una scheda arnia")
        cols = st.columns(len(ARNIE))
        for i, hive in enumerate(ARNIE):
            with cols[i]:
                if st.button(hive, key=f"overview_{hive}"):
                    open_hive(hive)

if page == "Scheda arnia":
    st.subheader("Scheda arnia")
    selected_hive = st.selectbox("Arnia", ARNIE, index=ARNIE.index(st.session_state["selected_hive"]) if st.session_state["selected_hive"] in ARNIE else 0)
    st.session_state["selected_hive"] = selected_hive
    hive_df = controlli[controlli["arnia"] == selected_hive].copy()
    if hive_df.empty:
        st.info("Nessun controllo per questa arnia.")
    else:
        hive_df["data_controllo_dt"] = pd.to_datetime(hive_df["data_controllo"], errors="coerce")
        hive_df = hive_df.sort_values("data_controllo_dt", ascending=False)
        last = hive_df.iloc[0]
        alerts, suggestions, next_visit = analyze_control(last)

        top1, top2 = st.columns([1.15, 1])
        with top1:
            st.markdown(
                f"""
                <div class="card">
                  <div style="font-weight:700;font-size:1.1rem;">{selected_hive}</div>
                  <div class="kpi">Ultimo controllo: {last['data_controllo']}</div>
                  <div class="kpi">Forza: {last['forza_colonia']}</div>
                  <div class="kpi">Telaini: {last['telaini_coperti']}</div>
                  <div class="kpi">Prossimo: {last['prossimo_controllo']}</div>
                  <div class="small">{last['note']}</div>
                </div>
                """,
                unsafe_allow_html=True
            )
        with top2:
            st.markdown("### Consiglio AI rapido")
            for level, text in alerts:
                css = "bee-ok" if level == "ok" else "bee-warn" if level == "warn" else "bee-danger"
                st.markdown(f'<div class="{css}">{text}</div>', unsafe_allow_html=True)
            for s in dict.fromkeys(suggestions):
                st.write(f"- {s}")
            st.info(f"Prossimo controllo suggerito: **{next_visit}**")

        st.markdown("### Storico controlli")
        view = hive_df[["data_controllo","forza_colonia","telaini_coperti","celle_reali","regina_nuova","melario_presente","melario_percento","api_nervose","prossimo_controllo","note"]].copy()
        st.dataframe(view, use_container_width=True, hide_index=True)

        hive_df["telaini_num"] = to_num(hive_df["telaini_coperti"])
        hive_df["melario_num"] = to_num(hive_df["melario_percento"])
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### Andamento telaini")
            chart1 = hive_df[["data_controllo_dt","telaini_num"]].dropna().set_index("data_controllo_dt").sort_index()
            if not chart1.empty:
                st.line_chart(chart1)
        with c2:
            st.markdown("### Andamento melario %")
            chart2 = hive_df[["data_controllo_dt","melario_num"]].dropna().set_index("data_controllo_dt").sort_index()
            if not chart2.empty:
                st.line_chart(chart2)

if page == "Overview magazzino":
    st.subheader("Overview magazzino")
    if acquisti.empty:
        st.info("Nessun articolo salvato in magazzino.")
    else:
        mag = acquisti.copy()
        mag["quantita_num"] = to_num(mag["quantita"])
        mag["prezzo_totale_num"] = to_num(mag["prezzo_totale"])
        st.dataframe(mag[["data_acquisto","categoria","prodotto","quantita","unita_misura","prezzo_totale","fornitore_sito"]].sort_values("data_acquisto", ascending=False), use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### Spesa per categoria")
            cat_spesa = mag.groupby("categoria", dropna=False)["prezzo_totale_num"].sum().reset_index()
            st.bar_chart(cat_spesa.set_index("categoria"))
        with c2:
            st.markdown("### Quantità per categoria")
            cat_qta = mag.groupby("categoria", dropna=False)["quantita_num"].sum().reset_index()
            st.bar_chart(cat_qta.set_index("categoria"))

if page == "Nuovo controllo":
    st.subheader("Nuovo controllo")
    with st.form("nuovo_controllo"):
        c1,c2 = st.columns(2)
        with c1:
            arnia = st.selectbox("Arnia", ARNIE, index=ARNIE.index(st.session_state["selected_hive"]))
            data_controllo = st.date_input("Data controllo", value=date.today(), format="DD/MM/YYYY")
            prossimo_controllo = st.date_input("Prossimo controllo", value=date.today(), format="DD/MM/YYYY")
            forza_colonia = st.selectbox("Forza colonia", FORZE)
            telaini_coperti = st.slider("Telaini coperti", 0, 10, 5)
        with c2:
            melario_presente = st.checkbox("Melario presente")
            melario_percento = st.slider("Melario pieno %", 0, 100, 0)
            covata_fresca = st.checkbox("Covata fresca")
            covata_opercolata = st.checkbox("Covata opercolata")
            celle_reali = st.checkbox("Celle reali")
            regina_vista = st.checkbox("Regina vista")
            regina_nuova = st.checkbox("Regina nuova / sospetta nuova")
            api_nervose = st.checkbox("Api nervose")
        nutrizione = st.selectbox("Nutrizione", NUTRIZIONI)
        foto = st.file_uploader("Foto", type=["jpg","jpeg","png"], accept_multiple_files=True)
        note = st.text_area("Note", height=180)
        if st.form_submit_button("Salva controllo"):
            img_json = save_uploaded_files(foto, arnia, str(data_controllo))
            row = {"id": uuid.uuid4().hex, "arnia": arnia, "data_controllo": str(data_controllo), "forza_colonia": forza_colonia,
                   "telaini_coperti": telaini_coperti, "covata_fresca": covata_fresca, "covata_opercolata": covata_opercolata,
                   "celle_reali": celle_reali, "regina_vista": regina_vista, "regina_nuova": regina_nuova, "melario_presente": melario_presente,
                   "melario_percento": melario_percento, "nutrizione": nutrizione, "api_nervose": api_nervose, "note": note,
                   "prossimo_controllo": str(prossimo_controllo), "foto": img_json}
            controlli = pd.concat([controlli, pd.DataFrame([row])], ignore_index=True)
            save_csv(controlli, FILES["controlli"])
            st.success("Controllo salvato.")

if page == "Consiglio AI":
    st.subheader("Consiglio AI")
    if controlli.empty:
        st.info("Salva almeno un controllo.")
    else:
        view = controlli.copy()
        view["data_controllo_dt"] = pd.to_datetime(view["data_controllo"], errors="coerce")
        view = view.sort_values("data_controllo_dt", ascending=False)
        options = [f"{r['arnia']} · {r['data_controllo']}" for _, r in view.iterrows()]
        selected = st.selectbox("Scegli il controllo", options)
        row = view.iloc[options.index(selected)]
        alerts, suggestions, next_visit = analyze_control(row)
        for level, text in alerts:
            css = "bee-ok" if level == "ok" else "bee-warn" if level == "warn" else "bee-danger"
            st.markdown(f'<div class="{css}">{text}</div>', unsafe_allow_html=True)
        for s in dict.fromkeys(suggestions):
            st.write(f"- {s}")
        st.info(f"Prossimo controllo suggerito: **{next_visit}**")

if page == "Magazzino":
    st.subheader("Magazzino")
    with st.form("magazzino"):
        c1,c2 = st.columns(2)
        with c1:
            data_acquisto = st.date_input("Data acquisto", value=date.today(), format="DD/MM/YYYY")
            categoria = st.selectbox("Categoria", MAG_CATEGORIE)
            prodotto = st.text_input("Prodotto")
            fornitore_sito = st.text_input("Sito / negozio / fornitore")
        with c2:
            quantita = st.number_input("Quantità", min_value=0.0, step=1.0)
            unita_misura = st.selectbox("Unità", ["pz", "kg", "confezioni", "litri", "altro"])
            prezzo_totale = st.number_input("Costo totale €", min_value=0.0, step=0.5, format="%.2f")
            note = st.text_area("Note", height=100)
        if st.form_submit_button("Salva articolo in magazzino"):
            row = {"id": uuid.uuid4().hex, "data_acquisto": str(data_acquisto), "categoria": categoria, "prodotto": prodotto, "quantita": quantita, "unita_misura": unita_misura, "prezzo_totale": prezzo_totale, "fornitore_sito": fornitore_sito, "note": note}
            acquisti = pd.concat([acquisti, pd.DataFrame([row])], ignore_index=True)
            save_csv(acquisti, FILES["acquisti"])
            st.success("Articolo salvato in magazzino.")
    if not acquisti.empty:
        st.dataframe(acquisti.sort_values("data_acquisto", ascending=False), use_container_width=True, hide_index=True)

if page == "Export Excel":
    st.subheader("Export Excel")
    if st.button("Prepara file Excel"):
        total_spese = float(to_num(acquisti.get("prezzo_totale", pd.Series(dtype=float))).sum()) if not acquisti.empty else 0.0
        bilancio = pd.DataFrame({"voce": ["Totale spese"], "valore": [total_spese]})
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            controlli.to_excel(writer, index=False, sheet_name="Controlli")
            acquisti.to_excel(writer, index=False, sheet_name="Magazzino")
            invasettato.to_excel(writer, index=False, sheet_name="Invasettato")
            vendite.to_excel(writer, index=False, sheet_name="Vendite")
            watchlist.to_excel(writer, index=False, sheet_name="Watchlist")
            offerte.to_excel(writer, index=False, sheet_name="Offerte")
            bilancio.to_excel(writer, index=False, sheet_name="Bilancio")
        st.download_button("Scarica Bee Honey.xlsx", data=output.getvalue(), file_name="Bee_Honey.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
