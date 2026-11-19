# Assegnazione dei Vangeli

Strumento che assegna un Vangelo a ogni collaboratore e prepara, per ciascuno, una cartella con le indicazioni e la traccia del contributo già compilata, più un riepilogo Excel.

Si può usare in due modi, con la **stessa logica** (il file `genera_cartelle.py`):

1. **Sito web**: `index.html` carica lo script Python nel browser (con Pyodide). Non serve installare nulla né avere un server.
2. **Riga di comando**: si lancia `genera_cartelle.py` sul proprio computer.

## Cosa fa

1. Legge i collaboratori dal CSV (nome, cognome, regione, telefono).
2. Legge il Word dei Vangeli e ricava per ogni giorno: intestazione, riferimento e testo.
3. **Esclude i giorni festivi**, cioè quelli con l'intestazione evidenziata (sfondo) in un colore stabilito, per default il verde `#b6d7a8`.
4. Assegna un Vangelo a ogni persona.
5. Crea una cartella `YYYYmmDD_Cognome_Nome` per ciascuna, dentro `AVVENTO-NATALE` o `QUARESIMA-PASQUA`.
6. In ogni cartella inserisce:
   - `Indicazioni_per_collaboratori.docx`, copia identica del modello (nessun marker sostituito);
   - `Traccia_contributo_Cognome_Nome.docx`, con i marker sostituiti da nome, giorno, riferimento e testo del Vangelo.
7. Crea `riepilogo.xlsx`.

## File di input

| File | Contenuto |
|---|---|
| CSV dei contatti | Intestazione con le colonne `nome`, `cognome`, `regione`, `telefono` (maiuscole e ordine indifferenti). Separatore `,` o `;`, codifica UTF-8. Il **telefono è facoltativo**: la colonna può mancare o avere celle vuote |
| Word dei Vangeli | Un blocco per giorno, vedi sotto |
| Indicazioni per collaboratori | Qualsiasi Word: viene copiato così com'è |
| Traccia contributo | Word con i 4 marker, vedi sotto |

**CSV da Excel:** File → Salva con nome → "CSV UTF-8 (delimitato dal separatore di elenco)". Da Fogli Google: File → Scarica → Valori separati da virgola. Conviene impostare la colonna telefono come testo, per non perdere zeri iniziali o il `+`.

**Regione ecclesiastica.** Viene calcolata dalla regione del CSV: Veneto, Trentino-Alto Adige e Friuli-Venezia Giulia diventano Triveneto, Abruzzo e Molise diventano Abruzzo-Molise, la Valle d'Aosta diventa Piemonte. Gli altri valori restano com'è scritto nel CSV.

### Marker della traccia

| Marker | Viene sostituito con |
|---|---|
| `[Nome + Cognome]` | nome e cognome della persona |
| `[GIORNO]` | intestazione del giorno, come scritta nel Word dei Vangeli |
| `[RIF VANGELO]` | es. "Dal Vangelo secondo Luca 1, 57-66" |
| `[TESTO VANGELO]` | testo del Vangelo, un paragrafo per ogni paragrafo del Vangelo |

Ogni marker deve essere l'unico testo del suo paragrafo, scritto esattamente così (maiuscole, spazi e parentesi quadre). La formattazione del marker viene mantenuta e il resto del documento non cambia. Modelli di esempio: cartella `docs/`.

### Struttura del Word dei Vangeli

- Ogni blocco ha: **intestazione** (una o due righe, con giorno e mese, es. `LUNEDÌ 20 DICEMBRE NOVENA`), **riferimento** (riga che inizia con `Dal Vangelo secondo …`, con capitolo e versetti anche sulla riga dopo) e **testo** (un paragrafo per riga, fino al giorno successivo).
- Se nell'intestazione manca il mese (`DOMENICA 19 IV AVVENTO`) viene preso dal giorno precedente.
- Le sezioni AVVENTO, NATALE, QUARESIMA e PASQUA vanno con lo stile "Titolo 2": decidono la macro categoria.
- Il primo paragrafo deve contenere l'anno liturgico nel formato `2027/2028`: novembre e dicembre cadono nel primo anno, gli altri mesi nel secondo.
- I **festivi** hanno uno *sfondo (ombreggiatura)* del colore stabilito sulle righe dell'intestazione.
- Se il giorno della settimana scritto non corrisponde alla data calcolata, viene segnalato un avviso.

## Uso 1: sito web

### File da pubblicare

```
index.html
genera_cartelle.py
docs/
├── Indicazioni_per_collaboratori_template.docx
└── Traccia_contributo_template.docx
```

Vanno bene host statici gratuiti come GitHub Pages, Netlify o Cloudflare Pages. Non caricare mai nel repository il CSV con i contatti né i Word reali.

### Come si usa la pagina

1. Carica i quattro file. L'icona ⓘ accanto a ciascuno spiega formato e marker.
2. Scegli il colore di sfondo che nel Word identifica i festivi (default verde), l'ordine di assegnazione (CSV o casuale, con seme per ripetere il risultato) e se usare i festivi come riserva.
3. **Anteprima assegnazione**: tabella scorrevole con data, persona, regione ecclesiastica, giorno liturgico, Vangelo e categoria. Le righe senza Vangelo sono in arancione e i festivi nel colore scelto.
4. **Genera**: scarica un unico file `vangeli_cartelle.zip` con le cartelle e `riepilogo.xlsx`.
5. **Aggiungere i telefoni dopo**: in fondo alla pagina carichi il riepilogo e un CSV con nome, cognome e telefono. I numeri vengono associati per nome e cognome e scarichi `riepilogo_con_telefoni.xlsx`.

I pulsanti "Scarica template" leggono i file dalla cartella `docs/`; se non li trovano, ne generano uno al volo. Se pubblichi usando la cartella `/docs` come radice del sito, in `index.html` imposta `TEMPLATE_DIR=''`.

### Come funziona e limiti

- Python gira nel browser con [Pyodide](https://pyodide.org) (versione 0.27.7 da jsDelivr). Al primo avvio scarica Python e le librerie `python-docx` e `openpyxl`: ci vuole circa un minuto, poi restano in cache. Serve quindi una connessione a internet.
- Nessun dato viene inviato a un server: i file restano nel browser.
- La pagina non funziona aprendo `index.html` direttamente dal computer, perché deve leggere `genera_cartelle.py`. Per provarla in locale, nella cartella lancia `python -m http.server` e apri `http://localhost:8000`.
- Durante la generazione la pagina può sembrare ferma per qualche decina di secondi.
- Ogni cartella pesa circa 2,3 MB per i font incorporati nei modelli Word. Con 131 persone lo ZIP supera i 200 MB e viene costruito in memoria: con poca RAM la generazione può fallire. In quel caso usa la riga di comando.

## Uso 2: riga di comando

Requisiti: Python 3.8 o successivo, `pip install python-docx openpyxl`.

```
python genera_cartelle.py
```

Per default cerca `contatti.csv`, `2027-2028_Tempo_Forti_C_revisionato.docx`, `Indicazioni_per_collaboratori.docx` e `Traccia_contributo.docx` nella cartella corrente e scrive in `output/`.

| Opzione | Effetto |
|---|---|
| `--csv`, `--vangeli`, `--indicazioni`, `--traccia` | Percorsi dei quattro file di input |
| `--output CARTELLA` | Cartella di destinazione (default `output`) |
| `--ordine csv` | Default: i Vangeli, in ordine cronologico, vanno alle persone nell'ordine del CSV |
| `--ordine casuale` | Le persone vengono mescolate prima dell'assegnazione |
| `--seed N` | Con `--ordine casuale`, rende il risultato ripetibile |
| `--includi-festivi` | Usa i festivi come riserva se i Vangeli non festivi non bastano |
| `--dry-run` | Stampa l'assegnazione senza creare file |

Il colore dei festivi, da riga di comando, si cambia nella costante `VERDE_FESTIVO` all'inizio dello script.

Esempi:

```
python genera_cartelle.py --dry-run
python genera_cartelle.py --ordine casuale --seed 7
python genera_cartelle.py --includi-festivi --output output_completo
```

## Struttura dell'output

```
output/
├── riepilogo.xlsx
├── AVVENTO-NATALE/
│   └── 20271223_Castaldi_Francesca/
│       ├── Indicazioni_per_collaboratori.docx
│       └── Traccia_contributo_Castaldi_Francesca.docx
└── QUARESIMA-PASQUA/
    └── ...
```

Negli ZIP del sito la struttura è identica. Nei nomi di cartella e file gli spazi interni diventano trattini (es. `De-Flaviis`) e i caratteri vietati vengono rimossi. I nomi scritti tutti in maiuscolo nel CSV vengono riportati a iniziali maiuscole.

### Riepilogo Excel

Colonne: Nome, Cognome, Regione, Regione ecclesiastica, Telefono, Data, Giorno liturgico, Vangelo (riferimento), Macro categoria, Cartella, Testo del Vangelo.

- Il telefono è salvato come testo.
- Le persone senza Vangelo hanno la riga evidenziata e "DA ASSEGNARE" nella colonna Data.
- Con l'uso dei festivi come riserva, le righe con Vangelo festivo sono evidenziate in verde chiaro.

## Note

- Con i file del 2027-2028 i Vangeli non festivi sono 108, i festivi 29 e i contatti 131: senza `--includi-festivi` (o l'opzione equivalente nella pagina) 23 persone restano "DA ASSEGNARE", senza cartella.
- Il campo `[GIORNO]` riporta l'intestazione così com'è nel Word dei Vangeli, senza l'anno.
- Se la traccia non contiene tutti e quattro i marker, la generazione si ferma con un errore.
- Rilanciando da riga di comando nella stessa cartella di output, i file esistenti vengono sovrascritti.
- Se la struttura del Word dei Vangeli cambia (nuovo ciclo, altro colore per i festivi), controlla le regole di lettura descritte sopra.
