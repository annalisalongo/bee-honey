import io
import json
import uuid
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import quote_plus

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
PESI_BARATTOLI = [125, 250, 500, 1000]

st.markdown("""
<style>
.main .block-container {padding-top:1rem;padding-bottom:2rem;max-width:1150px;}
div[data-testid="stMetric"] {background:#fff;border:1px solid #eadfbe;border-radius:18px;padding:8px 12px;}
div.stButton > button, .stDownloadButton button {width:100%;border-radius:14px;border:1px solid #d9b65d;background:#fff8e6;min-height:48px;font-weight:600;}
.bee-card {border:1px solid #eadfbe;background:linear-gradient(180deg,#fffdf7 0%,#fff9ea 100%);border-radius:18px;padding:16px;margin-bottom:12px;}
.bee-chip {display:inline-block;padding:4px 10px;border-radius:999px;border:1px solid #eadfbe;background:#fff;margin-right:6px;margin-bottom:6px;font-size:0.9rem;}
.bee-ok {background:#e8f7ee;border-left:6px solid #2c9b62;padding:12px;border-radius:12px;margin-bottom:10px;}
.bee-warn {background:#fff4df;border-left:6px solid #c98b00;padding:12px;border-radius:12px;margin-bottom:10px;}
.bee-danger {background:#fdecec;border-left:6px solid #c94b4b;padding:12px;border-radius:12px;margin-bottom:10px;}
.small-note {color:#6b604f;font-size:0.92rem;}
</style>
""", unsafe_allow_html=True)

def load_csv(path: Path, columns):
    if path.exists():
        df = pd.read_csv(path)
        for c in columns:
            if c not in df.columns:
                df[c] = ""
        return df[columns]
    return pd.DataFrame(columns=columns)

def save_csv(df, path: Path):
    df.to_csv(path, index=False)

def save_uploaded_files(files, folder_name: str, data_rif: str):
    saved_paths = []
    if not files:
        return ""
    safe_name = folder_name.lower().replace(" ", "_")
    folder = IMG_DIR / safe_name / data_rif
    folder.mkdir(parents=True, exist_ok=True)
    for f in files:
        ext = Path(f.name).suffix.lower() or ".jpg"
        filename = f"{uuid.uuid4().hex}{ext}"
        out = folder / filename
        with open(out, "wb") as fp:
            fp.write(f.getbuffer())
        saved_paths.append(str(out))
    return json.dumps(saved_paths, ensure_ascii=False)

def parse_images(value):
    if pd.isna(value) or not value:
        return []
    try:
        parsed = json.loads(value)
        return parsed if isinstance(parsed, list) else []
    except Exception:
        return []

def euro(value):
    try:
        return f"€ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "€ 0,00"

def bool_si_no(val):
    return "Sì" if str(val) == "True" else "No"

def to_num(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)

def analyze_control(row):
    alerts, suggestions, next_days = [], [], []
    forza = str(row.get("forza_colonia", ""))
    telaini = pd.to_numeric(pd.Series([row.get("telaini_coperti", 0)]), errors="coerce").fillna(0).iloc[0]
    covata_fresca = str(row.get("covata_fresca", "")) == "True"
    covata_opercolata = str(row.get("covata_opercolata", "")) == "True"
    celle_reali = str(row.get("celle_reali", "")) == "True"
    regina_vista = str(row.get("regina_vista", "")) == "True"
    regina_nuova = str(row.get("regina_nuova", "")) == "True"
    melario = str(row.get("melario_presente", "")) == "True"
    melario_pct = pd.to_numeric(pd.Series([row.get("melario_percento", 0)]), errors="coerce").fillna(0).iloc[0]
    api_nervose = str(row.get("api_nervose", "")) == "True"
    nutrizione = str(row.get("nutrizione", ""))

    if celle_reali and (forza in ["Forte", "Fortissima"] or telaini >= 8):
        alerts.append(("danger", "Rischio sciamatura alto: celle reali + colonia forte."))
        suggestions.append("Valuta divisione o controllo stretto delle celle reali.")
        next_days.append(3)
    if (not covata_fresca) and (not regina_vista) and (not regina_nuova):
        alerts.append(("warn", "Possibile problema regina: niente covata fresca e regina non vista."))
        suggestions.append("Controlla presenza di uova fresche e comportamento della colonia al prossimo giro.")
        next_days.append(5)
    if regina_nuova and not regina_vista:
        alerts.append(("ok", "Regina nuova o sospetta nuova: evita controlli troppo ravvicinati."))
        suggestions.append("Lascia tranquilla la colonia e fai un controllo mirato tra 7 e 10 giorni.")
        next_days.append(8)
    if melario and melario_pct >= 70:
        alerts.append(("warn", f"Melario al {int(melario_pct)}%: preparare un altro melario."))
        suggestions.append("Se il flusso è forte, aggiungi spazio prima che il melario sia pieno del tutto.")
        next_days.append(4)
    if forza == "Debole" and telaini <= 4:
        alerts.append(("warn", "Colonia debole: poco popolosa per ora."))
        suggestions.append("Non allargare troppo il nido. Tienila stretta e controlla scorte/regina.")
        next_days.append(7)
    if api_nervose:
        alerts.append(("warn", "Api nervose segnalate."))
        suggestions.append("Al prossimo controllo verifica meteo, presenza regina e manipolazione più rapida.")
        next_days.append(6)
    if (not covata_fresca) and (not covata_opercolata):
        alerts.append(("danger", "Nessuna covata segnata nel controllo."))
        suggestions.append("Verifica se è una pausa temporanea o un problema di regina/sciamatura.")
        next_days.append(4)
    if nutrizione in ["Candito", "Sciroppo"]:
        alerts.append(("ok", f"Nutrizione registrata: {nutrizione}."))
        suggestions.append("Controlla se la colonia consuma davvero la nutrizione e se sta ripartendo.")
        next_days.append(7)
    if (forza in ["Forte", "Fortissima"] or telaini >= 8) and not melario:
        alerts.append(("warn", "Colonia forte senza melario."))
        suggestions.append("Valuta se è il momento di dare spazio sopra per evitare congestione.")
        next_days.append(4)
    if not alerts:
        alerts.append(("ok", "Controllo nella norma."))
        suggestions.append("Continua con gestione regolare e monitoraggio settimanale.")
        next_days.append(7)
    next_visit = date.today() + timedelta(days=min(next_days))
    return alerts, suggestions, next_visit.strftime("%Y-%m-%d")

CONTROLLI_COLS = ["id","arnia","data_controllo","forza_colonia","telaini_coperti","covata_fresca","covata_opercolata","celle_reali","regina_vista","regina_nuova","melario_presente","melario_percento","nutrizione","api_nervose","note","prossimo_controllo","foto"]
ACQUISTI_COLS = ["id","data_acquisto","categoria","prodotto","quantita","unita_misura","prezzo_totale","fornitore_sito","note"]
INVASETTATO_COLS = ["id","data_invasettamento","peso_grammi","quantita_invasettata","lotto","note"]
VENDITE_COLS = ["id","data_vendita","peso_grammi","quantita_venduta","prezzo_unitario","canale_vendita","note"]
WATCHLIST_COLS = ["id","prodotto","categoria","prezzo_target","negozio_preferito","note"]
OFFERTE_COLS = ["id","data_offerta","prodotto","sito","url","prezzo","spedizione","prezzo_totale","prezzo_target","valutazione","note"]

controlli = load_csv(FILES["controlli"], CONTROLLI_COLS)
acquisti = load_csv(FILES["acquisti"], ACQUISTI_COLS)
invasettato = load_csv(FILES["invasettato"], INVASETTATO_COLS)
vendite = load_csv(FILES["vendite"], VENDITE_COLS)
watchlist = load_csv(FILES["watchlist"], WATCHLIST_COLS)
offerte = load_csv(FILES["offerte"], OFFERTE_COLS)

st.title("🐝 Bee Honey")
st.caption("Apiario, scorte, miele, bilancio, offerte e consigli automatici.")

menu = st.sidebar.radio("Menu", ["Home", "Arnie", "Nuovo controllo", "Consiglio AI", "Magazzino", "Miele invasettato", "Vendite", "Bilancio", "Bee Deals", "Export Excel"])

if menu == "Home":
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Arnie", len(ARNIE))
    c2.metric("Controlli", len(controlli))
    c3.metric("Acquisti", len(acquisti))
    c4.metric("Vendite", len(vendite))
    st.markdown("### Le tue arnie")
    today = pd.Timestamp(date.today())
    tmp = controlli.copy()
    if not tmp.empty:
        tmp["data_controllo_dt"] = pd.to_datetime(tmp["data_controllo"], errors="coerce")
        tmp["prossimo_controllo_dt"] = pd.to_datetime(tmp["prossimo_controllo"], errors="coerce")
    cols = st.columns(2)
    for i, arnia in enumerate(ARNIE):
        with cols[i % 2]:
            if tmp.empty or tmp[tmp["arnia"] == arnia].empty:
                st.markdown(f'<div class="bee-card"><h3>{arnia}</h3><div class="bee-chip">Nessun dato</div></div>', unsafe_allow_html=True)
            else:
                sub = tmp[tmp["arnia"] == arnia].sort_values("data_controllo_dt", ascending=False)
                last = sub.iloc[0]
                stato = "OK"
                box_class = "bee-ok"
                if pd.notna(last["prossimo_controllo_dt"]):
                    if last["prossimo_controllo_dt"] < today or last["prossimo_controllo_dt"] == today:
                        stato = "Urgente" if last["prossimo_controllo_dt"] < today else "Oggi"
                        box_class = "bee-warn"
                st.markdown(f'<div class="bee-card"><h3>{arnia}</h3><div class="bee-chip">Ultimo: {last["data_controllo"]}</div><div class="bee-chip">Prossimo: {last["prossimo_controllo"]}</div><div class="{box_class}"><b>Stato:</b> {stato}</div></div>', unsafe_allow_html=True)

if menu == "Arnie":
    arnia_filter = st.selectbox("Scegli arnia", ARNIE)
    view = controlli[controlli["arnia"] == arnia_filter].copy()
    if view.empty:
        st.info("Nessun controllo per questa arnia.")
    else:
        view["data_controllo_dt"] = pd.to_datetime(view["data_controllo"], errors="coerce")
        view = view.sort_values("data_controllo_dt", ascending=False)
        last = view.iloc[0]
        st.markdown(f'<div class="bee-card"><h3>{arnia_filter}</h3><div class="bee-chip">Ultimo controllo: {last["data_controllo"]}</div><div class="bee-chip">Prossimo: {last["prossimo_controllo"]}</div></div>', unsafe_allow_html=True)
        for _, row in view.iterrows():
            with st.expander(f'{row["data_controllo"]}'):
                st.write(f"**Forza colonia:** {row['forza_colonia']}")
                st.write(f"**Telaini coperti:** {row['telaini_coperti']}")
                st.write(f"**Covata fresca:** {bool_si_no(row['covata_fresca'])}")
                st.write(f"**Covata opercolata:** {bool_si_no(row['covata_opercolata'])}")
                st.write(f"**Celle reali:** {bool_si_no(row['celle_reali'])}")
                st.write(f"**Regina vista:** {bool_si_no(row['regina_vista'])}")
                st.write(f"**Melario:** {bool_si_no(row['melario_presente'])} ({row['melario_percento']}%)")
                st.write(f"**Api nervose:** {bool_si_no(row['api_nervose'])}")
                st.write(f"**Note:** {row['note'] or '—'}")

if menu == "Nuovo controllo":
    st.subheader("Nuovo controllo")
    with st.form("nuovo_controllo"):
        c1, c2 = st.columns(2)
        with c1:
            arnia = st.selectbox("Arnia", ARNIE)
            data_controllo = st.date_input("Data controllo", value=date.today(), format="DD/MM/YYYY")
            prossimo_controllo = st.date_input("Prossimo controllo", value=date.today(), format="DD/MM/YYYY")
            forza_colonia = st.selectbox("Forza colonia", ["Debole", "Media", "Forte", "Fortissima"])
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
        nutrizione = st.selectbox("Nutrizione", ["Nessuna", "Candito", "Sciroppo", "Altro"])
        foto = st.file_uploader("Foto", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        note = st.text_area("Note", height=180)
        submitted = st.form_submit_button("Salva controllo")
        if submitted:
            img_json = save_uploaded_files(foto, arnia, str(data_controllo))
            new_row = {"id": uuid.uuid4().hex, "arnia": arnia, "data_controllo": str(data_controllo), "forza_colonia": forza_colonia, "telaini_coperti": telaini_coperti, "covata_fresca": covata_fresca, "covata_opercolata": covata_opercolata, "celle_reali": celle_reali, "regina_vista": regina_vista, "regina_nuova": regina_nuova, "melario_presente": melario_presente, "melario_percento": melario_percento, "nutrizione": nutrizione, "api_nervose": api_nervose, "note": note, "prossimo_controllo": str(prossimo_controllo), "foto": img_json}
            controlli = pd.concat([controlli, pd.DataFrame([new_row])], ignore_index=True)
            save_csv(controlli, FILES["controlli"])
            st.success("Controllo salvato.")

if menu == "Consiglio AI":
    st.subheader("Consiglio AI")
    st.caption("Assistente intelligente basato sui dati che inserisci.")
    if controlli.empty:
        st.info("Salva almeno un controllo per avere consigli.")
    else:
        controlli_view = controlli.copy()
        controlli_view["data_controllo_dt"] = pd.to_datetime(controlli_view["data_controllo"], errors="coerce")
        controlli_view = controlli_view.sort_values("data_controllo_dt", ascending=False)
        options = [f"{row['arnia']} · {row['data_controllo']}" for _, row in controlli_view.iterrows()]
        selected = st.selectbox("Scegli il controllo da analizzare", options)
        row = controlli_view.iloc[options.index(selected)]
        alerts, suggestions, next_visit = analyze_control(row)
        st.markdown("### Analisi")
        for level, text in alerts:
            css = "bee-ok" if level == "ok" else "bee-warn" if level == "warn" else "bee-danger"
            st.markdown(f'<div class="{css}">{text}</div>', unsafe_allow_html=True)
        st.markdown("### Consigli")
        for i, suggestion in enumerate(dict.fromkeys(suggestions), start=1):
            st.write(f"{i}. {suggestion}")
        st.info(f"Prossimo controllo suggerito: **{next_visit}**")

if menu == "Magazzino":
    st.subheader("Magazzino e acquisti")
    with st.form("nuovo_acquisto"):
        c1, c2 = st.columns(2)
        with c1:
            data_acquisto = st.date_input("Data acquisto", value=date.today(), format="DD/MM/YYYY", key="data_acquisto")
            categoria = st.selectbox("Categoria", ["Fogli cerei", "Telaini", "Melari", "Barattoli", "Nutrizione", "Attrezzatura", "Altro"])
            prodotto = st.text_input("Prodotto")
            fornitore_sito = st.text_input("Sito / negozio / fornitore")
        with c2:
            quantita = st.number_input("Quantità", min_value=0.0, step=1.0)
            unita_misura = st.selectbox("Unità", ["pz", "kg", "confezioni", "litri", "altro"])
            prezzo_totale = st.number_input("Costo totale €", min_value=0.0, step=0.5, format="%.2f")
            note = st.text_area("Note", height=120)
        saved = st.form_submit_button("Salva acquisto")
        if saved:
            acquisti = pd.concat([acquisti, pd.DataFrame([{"id": uuid.uuid4().hex, "data_acquisto": str(data_acquisto), "categoria": categoria, "prodotto": prodotto, "quantita": quantita, "unita_misura": unita_misura, "prezzo_totale": prezzo_totale, "fornitore_sito": fornitore_sito, "note": note}])], ignore_index=True)
            save_csv(acquisti, FILES["acquisti"])
            st.success("Acquisto salvato.")
    if not acquisti.empty:
        st.dataframe(acquisti.sort_values("data_acquisto", ascending=False), use_container_width=True, hide_index=True)

if menu == "Miele invasettato":
    st.subheader("Carico barattoli invasettati")
    with st.form("carico_invasettato"):
        c1, c2 = st.columns(2)
        with c1:
            data_invasettamento = st.date_input("Data invasettamento", value=date.today(), format="DD/MM/YYYY")
            peso_grammi = st.selectbox("Peso barattolo (g)", PESI_BARATTOLI, key="peso_invas")
        with c2:
            quantita_invasettata = st.number_input("Quantità invasettata", min_value=0, step=1)
            lotto = st.text_input("Lotto")
        note = st.text_area("Note", height=120)
        saved = st.form_submit_button("Salva carico")
        if saved:
            invasettato = pd.concat([invasettato, pd.DataFrame([{"id": uuid.uuid4().hex, "data_invasettamento": str(data_invasettamento), "peso_grammi": peso_grammi, "quantita_invasettata": quantita_invasettata, "lotto": lotto, "note": note}])], ignore_index=True)
            save_csv(invasettato, FILES["invasettato"])
            st.success("Carico salvato.")
    if not invasettato.empty:
        st.dataframe(invasettato.sort_values("data_invasettamento", ascending=False), use_container_width=True, hide_index=True)

if menu == "Vendite":
    st.subheader("Vendite miele")
    with st.form("vendita"):
        c1, c2 = st.columns(2)
        with c1:
            data_vendita = st.date_input("Data vendita", value=date.today(), format="DD/MM/YYYY")
            peso_grammi = st.selectbox("Peso barattolo (g)", PESI_BARATTOLI)
            quantita_venduta = st.number_input("Quantità venduta", min_value=0, step=1)
        with c2:
            prezzo_unitario = st.number_input("Prezzo unitario €", min_value=0.0, step=0.5, format="%.2f")
            canale_vendita = st.selectbox("Canale vendita", ["Privati", "Amici", "Mercatino", "Regalo", "Altro"])
        note = st.text_area("Note", height=120)
        saved = st.form_submit_button("Salva vendita")
        if saved:
            vendite = pd.concat([vendite, pd.DataFrame([{"id": uuid.uuid4().hex, "data_vendita": str(data_vendita), "peso_grammi": peso_grammi, "quantita_venduta": quantita_venduta, "prezzo_unitario": prezzo_unitario, "canale_vendita": canale_vendita, "note": note}])], ignore_index=True)
            save_csv(vendite, FILES["vendite"])
            st.success("Vendita salvata.")
    if not vendite.empty:
        tmp = vendite.copy()
        tmp["quantita_venduta"] = to_num(tmp["quantita_venduta"])
        tmp["prezzo_unitario"] = to_num(tmp["prezzo_unitario"])
        tmp["incasso"] = tmp["quantita_venduta"] * tmp["prezzo_unitario"]
        st.dataframe(tmp.sort_values("data_vendita", ascending=False), use_container_width=True, hide_index=True)

if menu == "Bilancio":
    acquisti_tmp = acquisti.copy()
    vendite_tmp = vendite.copy()
    acquisti_tmp["prezzo_totale"] = to_num(acquisti_tmp.get("prezzo_totale", pd.Series(dtype=float)))
    vendite_tmp["quantita_venduta"] = to_num(vendite_tmp.get("quantita_venduta", pd.Series(dtype=float)))
    vendite_tmp["prezzo_unitario"] = to_num(vendite_tmp.get("prezzo_unitario", pd.Series(dtype=float)))
    vendite_tmp["incasso"] = vendite_tmp["quantita_venduta"] * vendite_tmp["prezzo_unitario"]
    spese = float(acquisti_tmp["prezzo_totale"].sum()) if not acquisti_tmp.empty else 0.0
    incassi = float(vendite_tmp["incasso"].sum()) if not vendite_tmp.empty else 0.0
    saldo = incassi - spese
    c1, c2, c3 = st.columns(3)
    c1.metric("Spese", euro(spese))
    c2.metric("Incassi", euro(incassi))
    c3.metric("Saldo", euro(saldo))

if menu == "Bee Deals":
    st.subheader("Bee Deals beta")
    tab1, tab2 = st.tabs(["Watchlist", "Nuova offerta"])
    with tab1:
        with st.form("watchlist_form"):
            prodotto = st.text_input("Prodotto da monitorare")
            prezzo_target = st.number_input("Prezzo target €", min_value=0.0, step=0.5, format="%.2f")
            save_watch = st.form_submit_button("Salva watchlist")
            if save_watch:
                watchlist = pd.concat([watchlist, pd.DataFrame([{"id": uuid.uuid4().hex, "prodotto": prodotto, "categoria": "", "prezzo_target": prezzo_target, "negozio_preferito": "", "note": ""}])], ignore_index=True)
                save_csv(watchlist, FILES["watchlist"])
                st.success("Prodotto aggiunto alla watchlist.")
        if not watchlist.empty:
            st.dataframe(watchlist, use_container_width=True, hide_index=True)
            for _, row in watchlist.iterrows():
                query = quote_plus(str(row["prodotto"]))
                st.markdown(f"- **{row['prodotto']}** · [Amazon](https://www.amazon.it/s?k={query}) · [eBay](https://www.ebay.it/sch/i.html?_nkw={query})")
    with tab2:
        st.info("Sezione offerte pronta per salvare occasioni manualmente nella versione successiva.")

if menu == "Export Excel":
    st.subheader("Export Excel")
    if st.button("Prepara file Excel"):
        total_spese = float(to_num(acquisti.get("prezzo_totale", pd.Series(dtype=float))).sum()) if not acquisti.empty else 0.0
        total_incassi = float((to_num(vendite.get("quantita_venduta", pd.Series(dtype=float))) * to_num(vendite.get("prezzo_unitario", pd.Series(dtype=float)))).sum()) if not vendite.empty else 0.0
        bilancio = pd.DataFrame({"voce": ["Totale spese", "Totale incassi", "Saldo"], "valore": [total_spese, total_incassi, total_incassi - total_spese]})
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
