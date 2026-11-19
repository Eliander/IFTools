#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera, per ogni collaboratore del CSV, una cartella YYYYmmDD_Cognome_Nome con:
  - Indicazioni_per_collaboratori.docx (copia identica)
  - Traccia_contributo_Cognome_Nome.docx (segnaposto sostituiti con nome, giorno, riferimento e testo del Vangelo)
e un file Excel di riepilogo.

Le cartelle sono divise in due macro categorie: AVVENTO-NATALE e QUARESIMA-PASQUA.
I giorni evidenziati in verde nel Word dei Vangeli (festivi) sono esclusi, a meno di --includi-festivi.

Requisiti:  pip install python-docx openpyxl

Esempi:
  python genera_cartelle.py                       # assegnazione in ordine di CSV
  python genera_cartelle.py --ordine casuale --seed 7
  python genera_cartelle.py --includi-festivi     # usa i festivi come riserva se mancano Vangeli
  python genera_cartelle.py --dry-run             # stampa solo l'assegnazione, non scrive nulla
"""
import argparse
import copy
import csv
import datetime as dt
import random
import re
import shutil
import sys
import unicodedata
from pathlib import Path

import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Configurazione di default (modificabile da riga di comando)
# --------------------------------------------------------------------------- #
CSV_DEFAULT = "contatti.csv"
VANGELI_DEFAULT = "2027-2028_Tempo_Forti_C_revisionato.docx"
INDICAZIONI_DEFAULT = "Indicazioni_per_collaboratori.docx"
TRACCIA_DEFAULT = "Traccia_contributo.docx"
OUTPUT_DEFAULT = "output"

VERDE_FESTIVO = "b6d7a8"  # colore di sfondo dei giorni festivi nel Word dei Vangeli

MESI = {
    "gennaio": 1, "febbraio": 2, "marzo": 3, "aprile": 4, "maggio": 5, "giugno": 6,
    "luglio": 7, "agosto": 8, "settembre": 9, "ottobre": 10, "novembre": 11, "dicembre": 12,
}
GIORNI_SETTIMANA = ["lunedi", "martedi", "mercoledi", "giovedi", "venerdi", "sabato", "domenica"]

# capitolo/versetti seguiti dall'inizio del testo, es. "Mt 23,1-12 In quel tempo, Gesù ..."
RE_CAPITOLO_E_TESTO = re.compile(
    r"^(\d+\s*,\s*[0-9ab.,\-–\s]*?\d[ab]?)\s+(?=[A-ZÀ-Ý«\[“\"])(.+)$"
)
# "Dal Vangelo secondo <Libro> [sigla] <capitolo, versetti ...>"
RE_TESTA_RIF = re.compile(
    r"^(Dal\s+(?:santo\s+)?Vangelo\s+secondo\s+[^\d]*?)\s*(?:\b[A-Z][a-z]{1,2}\.?\s+)?(\d.*)$", re.I
)

MACRO = {
    "AVVENTO": "AVVENTO-NATALE",
    "NATALE": "AVVENTO-NATALE",
    "QUARESIMA": "QUARESIMA-PASQUA",
    "PASQUA": "QUARESIMA-PASQUA",
}


# --------------------------------------------------------------------------- #
# Utilità
# --------------------------------------------------------------------------- #
def normalizza(s):
    """minuscolo, senza accenti né apostrofi: 'LUNEDI’' -> 'lunedi'."""
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[’'`´]", "", s).lower()


def pulisci_spazi(s):
    return re.sub(r"\s+", " ", s.replace("\t", " ")).strip()


def nome_file_sicuro(s):
    """Per i nomi di cartella/file: niente caratteri vietati, spazi -> '-'."""
    s = pulisci_spazi(s)
    s = re.sub(r'[\\/:*?"<>|]', "", s)
    return s.replace(" ", "-")


def e_verde(par):
    return any(e.get(qn("w:fill")) == VERDE_FESTIVO for e in par._p.iter(qn("w:shd")))


# --------------------------------------------------------------------------- #
# Lettura del Word dei Vangeli
# --------------------------------------------------------------------------- #
def leggi_vangeli(path):
    d = docx.Document(str(path))
    ps = d.paragraphs

    m = re.search(r"(\d{4})\s*/\s*(\d{4})", ps[0].text)
    anno_inizio = int(m.group(1)) if m else 2027

    # indici dei paragrafi "Dal Vangelo secondo ..."
    idx_vangelo = [i for i, p in enumerate(ps)
                   if re.match(r"\s*dal\s+(santo\s+)?vangelo", p.text.strip(), re.I)]

    # intestazione = paragrafi non vuoti (stile normale) subito sopra "Dal Vangelo"
    blocchi = []
    for i in idx_vangelo:
        hdr, j = [], i - 1
        while j >= 0 and ps[j].text.strip() and ps[j].style.name == "normal":
            hdr.insert(0, j)
            j -= 1
        blocchi.append({"i_rif": i, "hdr": hdr})

    # sezione liturgica (Heading 2) in cui cade ciascun blocco
    sezione_per_indice, corrente = {}, None
    for i, p in enumerate(ps):
        if p.style.name.lower().startswith("heading"):
            corrente = pulisci_spazi(p.text).upper()
        sezione_per_indice[i] = corrente

    vangeli, data_prec = [], None
    for k, b in enumerate(blocchi):
        hdr_txt = [pulisci_spazi(ps[j].text) for j in b["hdr"]]
        giorno = " – ".join(t for t in hdr_txt if t)
        festivo = any(e_verde(ps[j]) for j in b["hdr"])
        sezione = sezione_per_indice[b["i_rif"]]
        macro = MACRO.get(sezione, sezione)

        # riferimento: se la riga "Dal Vangelo secondo X" non contiene cifre, il capitolo è sulla riga dopo
        i = b["i_rif"]
        rif = pulisci_spazi(ps[i].text)
        corpo_da = i + 1
        incipit = None  # testo eventualmente presente sulla stessa riga del capitolo
        if not re.search(r"\d", rif):
            j = i + 1
            while j < len(ps) and not ps[j].text.strip():
                j += 1
            rif = f"{rif} {pulisci_spazi(ps[j].text)}"
            corpo_da = j + 1
        # pulizia: toglie la sigla ripetuta ("Matteo Mt 9,35-38" -> "Matteo 9,35-38") e separa
        # l'eventuale inizio del testo finito sulla stessa riga ("... 23,1-12 In quel tempo, ...")
        m3 = RE_TESTA_RIF.match(rif)
        if m3:
            testa, coda = m3.group(1).strip(), m3.group(2)
            m4 = RE_CAPITOLO_E_TESTO.match(coda)
            if m4:
                rif, incipit = f"{testa} {m4.group(1)}", m4.group(2)
            else:
                rif = f"{testa} {coda}"
        if len(rif) > 70:
            print(f"ATTENZIONE: riferimento sospetto (troppo lungo): {rif[:80]!r}", file=sys.stderr)

        # corpo: fino all'inizio del blocco successivo (o a un titolo di sezione)
        fine = blocchi[k + 1]["hdr"][0] if k + 1 < len(blocchi) else len(ps)
        corpo = [incipit] if incipit else []
        for j in range(corpo_da, fine):
            if ps[j].style.name.lower().startswith("heading"):
                break
            t = ps[j].text.strip()
            if t:
                corpo.append(t)

        # data
        testo_hdr = " ".join(hdr_txt)
        n = re.search(r"(?<!\d)(\d{1,2})(?!\d)", testo_hdr)
        if not n:
            raise ValueError(f"Giorno non trovato in: {giorno!r}")
        giorno_num = int(n.group(1))
        mese = next((v for nome, v in MESI.items() if nome in normalizza(testo_hdr)), None)
        if mese is None:  # es. "DOMENICA 19 IV AVVENTO": mese ereditato dal giorno precedente
            if data_prec is None:
                raise ValueError(f"Mese non trovato in: {giorno!r}")
            mese = data_prec.month + (1 if giorno_num < data_prec.day else 0)
        anno = anno_inizio if (macro == "AVVENTO-NATALE" and mese >= 11) else anno_inizio + 1
        data = dt.date(anno, mese, giorno_num)
        data_prec = data

        # controllo di coerenza giorno della settimana
        primo = normalizza(hdr_txt[0]).split()[0] if hdr_txt else ""
        if primo in GIORNI_SETTIMANA and GIORNI_SETTIMANA.index(primo) != data.weekday():
            print(f"ATTENZIONE: '{giorno}' -> {data} è un {GIORNI_SETTIMANA[data.weekday()]}",
                  file=sys.stderr)

        vangeli.append({
            "data": data, "giorno": giorno, "rif": rif, "testo": corpo,
            "festivo": festivo, "macro": macro, "sezione": sezione,
        })
    return vangeli


# --------------------------------------------------------------------------- #
# Lettura CSV
# --------------------------------------------------------------------------- #
def leggi_contatti(path):
    persone = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            r = {(k or "").strip().lower(): (v or "").strip() for k, v in r.items()}
            nome, cognome = pulisci_spazi(r.get("nome", "")), pulisci_spazi(r.get("cognome", ""))
            # nomi tutti in maiuscolo nel CSV -> iniziali maiuscole
            nome = nome.title() if nome.isupper() else nome
            cognome = cognome.title() if cognome.isupper() else cognome
            if not (nome or cognome):
                continue
            persone.append({
                "nome": nome, "cognome": cognome,
                "regione": pulisci_spazi(r.get("regione", "")),
                "telefono": pulisci_spazi(r.get("telefono", "")),
            })
    return persone


# --------------------------------------------------------------------------- #
# Compilazione della traccia
# --------------------------------------------------------------------------- #
def imposta_testo(par, testo):
    """Scrive il testo nel primo run (mantiene la formattazione) e rimuove gli altri."""
    runs = par.runs
    runs[0].text = testo
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def compila_traccia(template, destinazione, nome_completo, giorno, rif, testo_paragrafi):
    d = docx.Document(str(template))
    trovati = set()
    p_testo = None
    for p in d.paragraphs:
        t = p.text.strip()
        if t == "[Nome + Cognome]":
            imposta_testo(p, nome_completo); trovati.add("nome")
        elif t == "[GIORNO]":
            imposta_testo(p, giorno); trovati.add("giorno")
        elif t == "[RIF VANGELO]":
            imposta_testo(p, rif); trovati.add("rif")
        elif t == "[TESTO VANGELO]":
            p_testo = p; trovati.add("testo")
    if len(trovati) != 4:
        raise RuntimeError(f"Segnaposto mancanti nella traccia: {set(['nome','giorno','rif','testo']) - trovati}")

    # un paragrafo per ogni paragrafo del Vangelo (clonando quello del segnaposto)
    modello = copy.deepcopy(p_testo._p)
    imposta_testo(p_testo, testo_paragrafi[0])
    prev = p_testo._p
    for t in testo_paragrafi[1:]:
        nuovo = copy.deepcopy(modello)
        prev.addnext(nuovo)
        imposta_testo(Paragraph(nuovo, p_testo._parent), t)
        prev = nuovo
    d.save(str(destinazione))


# --------------------------------------------------------------------------- #
# Excel di riepilogo
# --------------------------------------------------------------------------- #
def scrivi_riepilogo(path, righe):
    wb = Workbook()
    ws = wb.active
    ws.title = "Riepilogo"
    intest = ["Nome", "Cognome", "Regione", "Telefono", "Data", "Giorno liturgico",
              "Vangelo (riferimento)", "Macro categoria", "Cartella", "Testo del Vangelo"]
    ws.append(intest)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for r in righe:
        v = r["vangelo"]
        ws.append([
            r["persona"]["nome"], r["persona"]["cognome"], r["persona"]["regione"],
            r["persona"]["telefono"],
            v["data"] if v else "DA ASSEGNARE",
            v["giorno"] if v else "", v["rif"] if v else "",
            v["macro"] if v else "", r.get("cartella", ""),
            "\n".join(v["testo"]) if v else "",
        ])
        riga = ws.max_row
        ws.cell(riga, 4).number_format = "@"  # telefono come testo
        if v:
            ws.cell(riga, 5).number_format = "DD/MM/YYYY"
            if v["festivo"]:
                for c in ws[riga]:
                    c.fill = PatternFill("solid", fgColor="E2EFDA")
        else:
            for c in ws[riga]:
                c.fill = PatternFill("solid", fgColor="FCE4D6")
    larghezze = [16, 20, 24, 18, 12, 44, 34, 18, 44, 80]
    for i, w in enumerate(larghezze, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(str(path))


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", default=CSV_DEFAULT)
    ap.add_argument("--vangeli", default=VANGELI_DEFAULT)
    ap.add_argument("--indicazioni", default=INDICAZIONI_DEFAULT)
    ap.add_argument("--traccia", default=TRACCIA_DEFAULT)
    ap.add_argument("--output", default=OUTPUT_DEFAULT)
    ap.add_argument("--ordine", choices=["csv", "casuale"], default="csv",
                    help="csv: i Vangeli sono assegnati in ordine cronologico alle persone nell'ordine del CSV (default); "
                         "casuale: le persone vengono mescolate prima dell'assegnazione")
    ap.add_argument("--seed", type=int, default=None, help="seme per --ordine casuale (riproducibilità)")
    ap.add_argument("--includi-festivi", action="store_true",
                    help="se i Vangeli non festivi non bastano, usa i festivi come riserva")
    ap.add_argument("--dry-run", action="store_true", help="stampa l'assegnazione senza creare file")
    a = ap.parse_args()

    vangeli = leggi_vangeli(a.vangeli)
    persone = leggi_contatti(a.csv)

    feriali = [v for v in vangeli if not v["festivo"]]
    festivi = [v for v in vangeli if v["festivo"]]
    disponibili = feriali + (festivi if a.includi_festivi else [])
    print(f"Vangeli: {len(vangeli)} totali, {len(feriali)} non festivi, {len(festivi)} festivi (esclusi"
          f"{' tranne se necessari' if a.includi_festivi else ''})")
    print(f"Persone: {len(persone)}")
    if len(persone) > len(disponibili):
        print(f"ATTENZIONE: {len(persone) - len(disponibili)} persone resteranno senza Vangelo "
              f"(usa --includi-festivi per usare anche i festivi).", file=sys.stderr)

    ordine = list(persone)
    if a.ordine == "casuale":
        random.Random(a.seed).shuffle(ordine)

    # un Vangelo per persona, in ordine cronologico
    righe = []
    for k, p in enumerate(ordine):
        righe.append({"persona": p, "vangelo": disponibili[k] if k < len(disponibili) else None})
    righe.sort(key=lambda r: (r["vangelo"] is None, r["vangelo"]["data"] if r["vangelo"] else dt.date.max,
                              r["persona"]["cognome"]))

    out = Path(a.output)
    for r in righe:
        v, p = r["vangelo"], r["persona"]
        if not v:
            continue
        cartella = f"{v['data']:%Y%m%d}_{nome_file_sicuro(p['cognome'])}_{nome_file_sicuro(p['nome'])}"
        r["cartella"] = f"{v['macro']}/{cartella}"

    if a.dry_run:
        for r in righe:
            v = r["vangelo"]
            print(f"{r.get('cartella', '-- DA ASSEGNARE --'):<60} {v['rif'] if v else ''}")
        return

    out.mkdir(parents=True, exist_ok=True)
    for r in righe:
        v, p = r["vangelo"], r["persona"]
        if not v:
            continue
        cart = out / r["cartella"]
        cart.mkdir(parents=True, exist_ok=True)
        shutil.copy2(a.indicazioni, cart / Path(a.indicazioni).name)
        nome_file = f"Traccia_contributo_{nome_file_sicuro(p['cognome'])}_{nome_file_sicuro(p['nome'])}.docx"
        compila_traccia(a.traccia, cart / nome_file, f"{p['nome']} {p['cognome']}",
                        v["giorno"], v["rif"], v["testo"])

    scrivi_riepilogo(out / "riepilogo.xlsx", righe)
    n = sum(1 for r in righe if r["vangelo"])
    print(f"Fatto: {n} cartelle create in '{out}', riepilogo in '{out / 'riepilogo.xlsx'}'.")


if __name__ == "__main__":
    main()
