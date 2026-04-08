# Bee Honey

Versione completa con:
- arnie Bee Bianca, Bee Verdina, Bee Gialla, Bee Verde
- controlli con foto
- magazzino e acquisti
- miele invasettato
- vendite
- grafici semplici
- export Excel unico
- Bee Deals beta con watchlist e confronto offerte

## Avvio locale
```bash
cd ~/Downloads/bee_honey_complete_v2
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m streamlit run app.py
```

## Da telefono
```bash
python -m streamlit run app.py --server.address 0.0.0.0
```

Poi apri:
`http://IP_DEL_MAC:8501`

## Nota su Bee Deals
La parte offerte è una beta intelligente:
- ti aiuta a tenere una watchlist
- classifica le offerte che trovi
- genera link rapidi Amazon/eBay

Per un monitoraggio automatico reale dei prezzi serve un passaggio successivo.
