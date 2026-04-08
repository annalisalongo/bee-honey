import io
import json
import uuid
from datetime import date
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
:root {
  --gold: #d9b65d;
  --card: #ffffff;
  --line: #eadfbe;
  --ok: #e8f7ee;
  --warn: #fff4df;
}
.main .block-container {
  padding-top: 1rem;
  padding-bottom: 2rem;
  max-width: 1150px;
}
div[data-testid="stMetric"] {
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 18px;
  padding: 8px 12px;
}
div.stButton > button, .stDownloadButton button {
  width: 100%;
  border-radius: 14px;
  border: 1px solid var(--gold);
  background: #fff8e6;
  min-height: 48px;
  font-weight: 600;
}
.bee-card {
  border: 1px solid var(--line);
  background: linear-gradient(180deg,#fffdf7 0%,#fff9ea 100%);
  border-radius: 18px;
  padding: 16px;
  margin-bottom: 12px;
}
.bee-chip {
  display: inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  border: 1px solid var(--line);
  background: #fff;
  margin-right: 6px;
  margin-bottom: 6px;
  font-size: 0.9rem;
}
.bee-ok {
  background: var(--ok);
  border-left: 6px solid #2c9b62;
  padding: 12px;
  border-radius: 12px;
  margin-bottom: 10px;
}
.bee-warn {
  background: var(--warn);
  border-left: 6px solid #c98b00;
  padding: 12px;
  border-radius: 12px;
  margin-bottom: 10px;
}
.small-note {
  color: #6b604f;
  font-size: 0.92rem;
}
</style>
""", unsafe_allow_html=True)


def load_csv(path: Path, columns: list[str]) -> pd.DataFrame:
    if path.exists():
        df = pd.read_csv(path)
        for c in columns:
            if c not in df.columns:
                df[c] = ""
        return df[columns]
    return pd.DataFrame(columns=columns)


def save_csv(df: pd.DataFrame, path: Path):
    df.to_csv(path, index=False)


def save_uploaded_files(files, folder_name: str, data_rif: str) -> str:
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


def euro(value) -> str:
    try:
        return f"€ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "€ 0,00"


def bool_si_no(val) -> str:
    return "Sì" if str(val) == "True" else "No"


def to_num(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


CONTROLLI_COLS = [
    "id","arnia","data_controllo","forza_colonia","telaini_coperti","covata_fresca",
    "covata_opercolata","celle_reali","regina_vista","regina_nuova","melario_presente",
    "melario_percento","nutrizione","api_nervose","note","prossimo_controllo","foto"
]
ACQUISTI_COLS = [
    "id","data_acquisto","categoria","prodotto","quantita","unita_misura",
    "prezzo_totale","fornitore_sito","note"
]
INVASETTATO_COLS = [
    "id","data_invasettamento","peso_grammi","quantita_invasettata","lotto","note"
]
VENDITE_COLS = [
    "id","data_vendita","peso_grammi","quantita_venduta","prezzo_unitario","canale_vendita","note"
]
WATCHLIST_COLS = [
    "id","prodotto","categoria","prezzo_target","negozio_preferito","note"
]
OFFERTE_COLS = [
    "id","data_offerta","prodotto","sito","url","prezzo","spedizione","prezzo_totale","prezzo_target","valutazione","note"
]

controlli = load_csv(FILES["controlli"], CONTROLLI_COLS)
acquisti = load_csv(FILES["acquisti"], ACQUISTI_COLS)
invasettato = load_csv(FILES["invasettato"], INVASETTATO_COLS)
vendite = load_csv(FILES["vendite"], VENDITE_COLS)
watchlist = load_csv(FILES["watchlist"], WATCHLIST_COLS)
offerte = load_csv(FILES["offerte"], OFFERTE_COLS)

st.title("🐝 Bee Honey")
st.caption("Apiario, scorte, miele, bilancio ed offerte in un'unica app.")

menu = st.sidebar.radio(
    "Menu",
    ["Home", "Arnie", "Nuovo controllo", "Magazzino", "Miele invasettato", "Vendite", "Bilancio", "Bee Deals", "Export Excel"]
)

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
                st.markdown(
                    f'<div class="bee-card"><h3>{arnia}</h3><div class="bee-chip">Nessun dato</div></div>',
                    unsafe_allow_html=True
                )
            else:
                sub = tmp[tmp["arnia"] == arnia].sort_values("data_controllo_dt", ascending=False)
                last = sub.iloc[0]
                stato = "OK"
                box_class = "bee-ok"
                if pd.notna(last["prossimo_controllo_dt"]):
                    if last["prossimo_controllo_dt"] < today:
                        stato = "Urgente"
                        box_class = "bee-warn"
                    elif last["prossimo_controllo_dt"] == today:
                        stato = "Oggi"
                        box_class = "bee-warn"
                html = f"""<div class="bee-card">
                  <h3>{arnia}</h3>
                  <div class="bee-chip">Ultimo: {last["data_controllo"]}</div>
                  <div class="bee-chip">Prossimo: {last["prossimo_controllo"]}</div>
                  <div class="{box_class}"><b>Stato:</b> {stato}</div>
                </div>"""
                st.markdown(html, unsafe_allow_html=True)

    st.markdown("### Oggi cosa guardare")
    if not tmp.empty:
        due = tmp[tmp["prossimo_controllo_dt"].notna() & (tmp["prossimo_controllo_dt"] <= today)]
        if due.empty:
            st.markdown('<div class="bee-ok">Nessun controllo urgente oggi.</div>', unsafe_allow_html=True)
        else:
            for _, row in due.sort_values("prossimo_controllo_dt").iterrows():
                note = row["note"] if isinstance(row["note"], str) else ""
                short_note = note[:140] + ("..." if len(note) > 140 else "")
                html = f'<div class="bee-warn"><b>{row["arnia"]}</b> da controllare il {row["prossimo_controllo"]}<br><span class="small-note">{short_note}</span></div>'
                st.markdown(html, unsafe_allow_html=True)

if menu == "Arnie":
    arnia_filter = st.selectbox("Scegli arnia", ARNIE)
    view = controlli[controlli["arnia"] == arnia_filter].copy()
    if view.empty:
        st.info("Nessun controllo per questa arnia.")
    else:
        view["data_controllo_dt"] = pd.to_datetime(view["data_controllo"], errors="coerce")
        view = view.sort_values("data_controllo_dt", ascending=False)
        last = view.iloc[0]
        html = f'<div class="bee-card"><h3>{arnia_filter}</h3><div class="bee-chip">Ultimo controllo: {last["data_controllo"]}</div><div class="bee-chip">Prossimo: {last["prossimo_controllo"]}</div></div>'
        st.markdown(html, unsafe_allow_html=True)
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
                imgs = parse_images(row["foto"])
                existing = [p for p in imgs if Path(p).exists()]
                if existing:
                    st.image(existing, width=220)

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
            new_row = {
                "id": uuid.uuid4().hex, "arnia": arnia, "data_controllo": str(data_controllo),
                "forza_colonia": forza_colonia, "telaini_coperti": telaini_coperti,
                "covata_fresca": covata_fresca, "covata_opercolata": covata_opercolata,
                "celle_reali": celle_reali, "regina_vista": regina_vista, "regina_nuova": regina_nuova,
                "melario_presente": melario_presente, "melario_percento": melario_percento,
                "nutrizione": nutrizione, "api_nervose": api_nervose, "note": note,
                "prossimo_controllo": str(prossimo_controllo), "foto": img_json
            }
            controlli = pd.concat([controlli, pd.DataFrame([new_row])], ignore_index=True)
            save_csv(controlli, FILES["controlli"])
            st.success("Controllo salvato.")

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
            new_row = {
                "id": uuid.uuid4().hex, "data_acquisto": str(data_acquisto), "categoria": categoria,
                "prodotto": prodotto, "quantita": quantita, "unita_misura": unita_misura,
                "prezzo_totale": prezzo_totale, "fornitore_sito": fornitore_sito, "note": note
            }
            acquisti = pd.concat([acquisti, pd.DataFrame([new_row])], ignore_index=True)
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
            new_row = {
                "id": uuid.uuid4().hex, "data_invasettamento": str(data_invasettamento),
                "peso_grammi": peso_grammi, "quantita_invasettata": quantita_invasettata,
                "lotto": lotto, "note": note
            }
            invasettato = pd.concat([invasettato, pd.DataFrame([new_row])], ignore_index=True)
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
            new_row = {
                "id": uuid.uuid4().hex, "data_vendita": str(data_vendita), "peso_grammi": peso_grammi,
                "quantita_venduta": quantita_venduta, "prezzo_unitario": prezzo_unitario,
                "canale_vendita": canale_vendita, "note": note
            }
            vendite = pd.concat([vendite, pd.DataFrame([new_row])], ignore_index=True)
            save_csv(vendite, FILES["vendite"])
            st.success("Vendita salvata.")

    if not vendite.empty:
        tmp = vendite.copy()
        tmp["quantita_venduta"] = to_num(tmp["quantita_venduta"])
        tmp["prezzo_unitario"] = to_num(tmp["prezzo_unitario"])
        tmp["incasso"] = tmp["quantita_venduta"] * tmp["prezzo_unitario"]
        st.dataframe(tmp.sort_values("data_vendita", ascending=False), use_container_width=True, hide_index=True)

        if not invasettato.empty:
            inv = invasettato.copy()
            inv["quantita_invasettata"] = to_num(inv["quantita_invasettata"])
            residuo = inv.groupby("peso_grammi")["quantita_invasettata"].sum().subtract(
                tmp.groupby("peso_grammi")["quantita_venduta"].sum(), fill_value=0
            ).reset_index()
            residuo.columns = ["peso_grammi", "barattoli_residui"]
            st.markdown("### Residuo in magazzino")
            st.dataframe(residuo, use_container_width=True, hide_index=True)

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

    cc1, cc2 = st.columns(2)
    with cc1:
        if not acquisti_tmp.empty:
            st.markdown("### Spese per categoria")
            spese_cat = acquisti_tmp.groupby("categoria", dropna=False)["prezzo_totale"].sum().reset_index()
            st.bar_chart(spese_cat.set_index("categoria"))
    with cc2:
        if not vendite_tmp.empty:
            st.markdown("### Ricavi per formato")
            ricavi = vendite_tmp.groupby("peso_grammi", dropna=False)["incasso"].sum().reset_index()
            st.bar_chart(ricavi.set_index("peso_grammi"))

    c3, c4 = st.columns(2)
    with c3:
        if not acquisti_tmp.empty:
            acquisti_time = acquisti_tmp.copy()
            acquisti_time["data_acquisto"] = pd.to_datetime(acquisti_time["data_acquisto"], errors="coerce")
            acquisti_time = acquisti_time.dropna(subset=["data_acquisto"]).sort_values("data_acquisto")
            if not acquisti_time.empty:
                serie_spese = acquisti_time.groupby("data_acquisto")["prezzo_totale"].sum()
                st.markdown("### Spese nel tempo")
                st.line_chart(serie_spese)
    with c4:
        if not vendite_tmp.empty:
            vendite_time = vendite_tmp.copy()
            vendite_time["data_vendita"] = pd.to_datetime(vendite_time["data_vendita"], errors="coerce")
            vendite_time = vendite_time.dropna(subset=["data_vendita"]).sort_values("data_vendita")
            if not vendite_time.empty:
                serie_incassi = vendite_time.groupby("data_vendita")["incasso"].sum()
                st.markdown("### Incassi nel tempo")
                st.line_chart(serie_incassi)

if menu == "Bee Deals":
    st.subheader("Bee Deals beta")
    st.caption("Per ora è una sezione smart per confrontare offerte che trovi tu e tenere una watchlist con prezzi target.")

    tab1, tab2, tab3 = st.tabs(["Watchlist", "Nuova offerta", "Offerte salvate"])

    with tab1:
        with st.form("watchlist_form"):
            c1, c2 = st.columns(2)
            with c1:
                prodotto = st.text_input("Prodotto da monitorare")
                categoria = st.selectbox("Categoria", ["Fogli cerei", "Telaini", "Melari", "Barattoli", "Nutrizione", "Attrezzatura", "Altro"], key="wl_cat")
            with c2:
                prezzo_target = st.number_input("Prezzo target €", min_value=0.0, step=0.5, format="%.2f")
                negozio_preferito = st.text_input("Negozio preferito (facoltativo)")
            note = st.text_area("Note")
            save_watch = st.form_submit_button("Salva watchlist")
            if save_watch:
                new_row = {
                    "id": uuid.uuid4().hex,
                    "prodotto": prodotto,
                    "categoria": categoria,
                    "prezzo_target": prezzo_target,
                    "negozio_preferito": negozio_preferito,
                    "note": note,
                }
                watchlist = pd.concat([watchlist, pd.DataFrame([new_row])], ignore_index=True)
                save_csv(watchlist, FILES["watchlist"])
                st.success("Prodotto aggiunto alla watchlist.")

        if not watchlist.empty:
            st.dataframe(watchlist, use_container_width=True, hide_index=True)
            st.markdown("### Ricerca rapida")
            for _, row in watchlist.iterrows():
                query = quote_plus(str(row["prodotto"]))
                amazon = f"https://www.amazon.it/s?k={query}"
                ebay = f"https://www.ebay.it/sch/i.html?_nkw={query}"
                st.markdown(f"- **{row['prodotto']}** · [Amazon]({amazon}) · [eBay]({ebay})")

    with tab2:
        with st.form("offerta_form"):
            c1, c2 = st.columns(2)
            with c1:
                data_offerta = st.date_input("Data offerta", value=date.today(), format="DD/MM/YYYY")
                prodotto = st.text_input("Prodotto", key="off_prod")
                sito = st.selectbox("Sito", ["Amazon", "eBay", "Altro sito"], key="off_site")
                url = st.text_input("Link prodotto")
            with c2:
                prezzo = st.number_input("Prezzo €", min_value=0.0, step=0.5, format="%.2f")
                spedizione = st.number_input("Spedizione €", min_value=0.0, step=0.5, format="%.2f")
                prezzo_target = st.number_input("Prezzo target €", min_value=0.0, step=0.5, format="%.2f")
                note = st.text_area("Note")
            save_offer = st.form_submit_button("Analizza e salva")
            if save_offer:
                prezzo_totale = float(prezzo) + float(spedizione)
                if prezzo_target <= 0:
                    valutazione = "Da valutare"
                elif prezzo_totale <= prezzo_target * 0.8:
                    valutazione = "Super conveniente"
                elif prezzo_totale <= prezzo_target:
                    valutazione = "Conveniente"
                else:
                    valutazione = "Normale"
                new_row = {
                    "id": uuid.uuid4().hex,
                    "data_offerta": str(data_offerta),
                    "prodotto": prodotto,
                    "sito": sito,
                    "url": url,
                    "prezzo": prezzo,
                    "spedizione": spedizione,
                    "prezzo_totale": prezzo_totale,
                    "prezzo_target": prezzo_target,
                    "valutazione": valutazione,
                    "note": note,
                }
                offerte = pd.concat([offerte, pd.DataFrame([new_row])], ignore_index=True)
                save_csv(offerte, FILES["offerte"])
                if valutazione == "Super conveniente":
                    st.success(f"{valutazione}: totale {euro(prezzo_totale)}")
                elif valutazione == "Conveniente":
                    st.info(f"{valutazione}: totale {euro(prezzo_totale)}")
                else:
                    st.warning(f"{valutazione}: totale {euro(prezzo_totale)}")

    with tab3:
        if offerte.empty:
            st.info("Nessuna offerta salvata.")
        else:
            offerte_view = offerte.copy()
            offerte_view["prezzo_totale"] = to_num(offerte_view["prezzo_totale"])
            st.dataframe(offerte_view.sort_values("data_offerta", ascending=False), use_container_width=True, hide_index=True)
            top = offerte_view[offerte_view["valutazione"].isin(["Super conveniente", "Conveniente"])]
            if not top.empty:
                st.markdown("### Migliori occasioni")
                for _, row in top.sort_values("data_offerta", ascending=False).iterrows():
                    html = f'<div class="bee-ok"><b>{row["prodotto"]}</b> · {row["sito"]} · {euro(row["prezzo_totale"])} · {row["valutazione"]}</div>'
                    st.markdown(html, unsafe_allow_html=True)

if menu == "Export Excel":
    st.subheader("Export Excel")
    st.caption("Un file unico con più fogli, pronto per Excel.")
    if st.button("Prepara file Excel"):
        total_spese = float(to_num(acquisti.get("prezzo_totale", pd.Series(dtype=float))).sum()) if not acquisti.empty else 0.0
        total_incassi = float((to_num(vendite.get("quantita_venduta", pd.Series(dtype=float))) * to_num(vendite.get("prezzo_unitario", pd.Series(dtype=float)))).sum()) if not vendite.empty else 0.0
        bilancio = pd.DataFrame(
            {"voce": ["Totale spese", "Totale incassi", "Saldo"], "valore": [total_spese, total_incassi, total_incassi - total_spese]}
        )
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            controlli.to_excel(writer, index=False, sheet_name="Controlli")
            acquisti.to_excel(writer, index=False, sheet_name="Magazzino")
            invasettato.to_excel(writer, index=False, sheet_name="Invasettato")
            vendite.to_excel(writer, index=False, sheet_name="Vendite")
            watchlist.to_excel(writer, index=False, sheet_name="Watchlist")
            offerte.to_excel(writer, index=False, sheet_name="Offerte")
            bilancio.to_excel(writer, index=False, sheet_name="Bilancio")
        st.download_button(
            "Scarica Bee Honey.xlsx",
            data=output.getvalue(),
            file_name="Bee_Honey.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
