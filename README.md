# genera_cartelle.py

Script Python che assegna un Vangelo a ogni collaboratore e prepara, per ciascuno, una cartella con le indicazioni e la traccia del contributo già compilata. Produce anche un Excel di riepilogo.

## Cosa fa

1. Legge l'elenco dei collaboratori dal CSV (nome, cognome, regione, telefono).
2. Legge il Word dei Vangeli e ricava per ogni giorno: intestazione, riferimento e testo.
3. **Esclude i giorni festivi**, cioè quelli con l'intestazione evidenziata in verde nel Word.
4. Assegna un Vangelo a ogni persona.
5. Crea una cartella `YYYYmmDD_Cognome_Nome` per ciascuna persona, dentro una delle due macro categorie `AVVENTO-NATALE` e `QUARESIMA-PASQUA`.
6. In ogni cartella inserisce:
   - `Indicazioni_per_collaboratori.docx`, copia identica del modello;
   - `Traccia_contributo_Cognome_Nome.docx`, con nome e cognome, giorno, riferimento e testo del Vangelo al posto dei segnaposto.
7. Crea `riepilogo.xlsx` nella cartella di output.

## Requisiti

- Python 3.8 o successivo
- Librerie: `python-docx` e `openpyxl`

```
pip install python-docx openpyxl
```

## File di input

Per default lo script li cerca nella cartella in cui viene lanciato.

| File | Contenuto |
|---|---|
| `contatti.csv` | Colonne `NOME`, `COGNOME`, `REGIONE`, `Telefono` (UTF-8, separatore virgola) |
| `2027-2028_Tempo_Forti_C_revisionato.docx` | Elenco dei Vangeli |
| `Indicazioni_per_collaboratori.docx` | Istruzioni, copiate così come sono in ogni cartella |
| `Traccia_contributo.docx` | Modello con i segnaposto `[Nome + Cognome]`, `[GIORNO]`, `[RIF VANGELO]`, `[TESTO VANGELO]` |

## Uso

```
python genera_cartelle.py
```

Il risultato è nella cartella `output/`.

### Opzioni

| Opzione | Effetto |
|---|---|
| `--csv`, `--vangeli`, `--indicazioni`, `--traccia` | Percorsi dei quattro file di input, se diversi dai nomi di default |
| `--output CARTELLA` | Cartella di destinazione (default `output`) |
| `--ordine csv` | Default. I Vangeli, in ordine cronologico, vanno alle persone nell'ordine del CSV |
| `--ordine casuale` | Le persone vengono mescolate prima dell'assegnazione |
| `--seed N` | Con `--ordine casuale`, rende il risultato ripetibile |
| `--includi-festivi` | Se i Vangeli non festivi non bastano, usa i festivi come riserva |
| `--dry-run` | Stampa l'assegnazione senza creare file |

Esempi:

```
python genera_cartelle.py --dry-run
python genera_cartelle.py --ordine casuale --seed 7
python genera_cartelle.py --includi-festivi --output output_completo
```

Consiglio: prima un `--dry-run` per controllare l'assegnazione, poi la generazione vera.

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

Nei nomi di cartella e file gli spazi interni diventano trattini (es. `De-Flaviis`) e i caratteri vietati dal file system vengono rimossi.

### Riepilogo Excel

Colonne: Nome, Cognome, Regione, Telefono, Data, Giorno liturgico, Vangelo (riferimento), Macro categoria, Cartella, Testo del Vangelo.

- Il telefono è salvato come testo, così non perde zeri o prefissi.
- Le persone senza Vangelo hanno la riga evidenziata e "DA ASSEGNARE" nella colonna Data.
- Se si usa `--includi-festivi`, le righe con Vangelo festivo sono evidenziate in verde chiaro.

## Come viene letto il Word dei Vangeli

- Ogni blocco inizia con un paragrafo "Dal Vangelo secondo ...". L'intestazione del giorno è il gruppo di paragrafi non vuoti subito sopra.
- **Festivo** = sfondo verde (`b6d7a8`) su almeno una riga dell'intestazione.
- **Macro categoria** = ultimo titolo di sezione (Heading 2) prima del blocco: AVVENTO e NATALE diventano `AVVENTO-NATALE`, QUARESIMA e PASQUA diventano `QUARESIMA-PASQUA`.
- **Anno**: ricavato dal titolo "2027/2028". Novembre e dicembre cadono nel primo anno, tutto il resto nel secondo.
- **Mese mancante** (es. "DOMENICA 19 IV AVVENTO"): ereditato dal giorno precedente.
- **Controllo di coerenza**: se il giorno della settimana scritto non corrisponde alla data calcolata, lo script stampa un avviso.
- **Riferimento**: la sigla ripetuta ("Matteo Mt 9,35-38") viene eliminata e, se il testo inizia sulla stessa riga del riferimento, viene separato.
- Il testo del Vangelo viene inserito come testo semplice, un paragrafo per ogni paragrafo del Word.

## Note e limiti

- Il campo `[GIORNO]` riporta l'intestazione così com'è nel Word dei Vangeli (es. "LUNEDI’ 20 DICEMBRE NOVENA"), senza l'anno.
- Se i Vangeli non festivi sono meno delle persone, le ultime nell'ordine di assegnazione restano "DA ASSEGNARE", senza cartella. Con i file attuali: 108 Vangeli non festivi, 29 festivi, 131 persone.
- I modelli Word hanno i font incorporati, quindi ogni cartella pesa circa 2,3 MB.
- Se la struttura del Word dei Vangeli cambia (nuovo ciclo, altro colore per i festivi), controllare le costanti all'inizio dello script, in particolare `VERDE_FESTIVO`.
- Se la traccia non contiene tutti e quattro i segnaposto, lo script si ferma con un errore.
- Rilanciando lo script con la stessa cartella di output, i file esistenti vengono sovrascritti.
