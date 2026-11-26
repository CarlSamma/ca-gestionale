# -*- coding: utf-8 -*-
"""
genera_mora.py - Genera la lettera di messa in mora di Como Acqua s.r.l.
dal modello dello Studio (template/modello_messa_in_mora.docx) e da un file
dati_mora.json (prodotto da estrai_pratica.py, modificabile a mano).

Uso:
    py genera_mora.py <cartella>/dati_mora.json [--doc] [--pdf] [--no-gestionale]

Uscite nella cartella della pratica:
    <output>              (.docx, default "messa in mora_<debitore>.docx")
    <output>              (.doc  con --doc,  .pdf con --pdf)  via Word COM

Inoltre aggiorna/crea L:\\PROGETTI\\COMOACQUA\\GESTIONALE\\gestionale_ca.json
aggiungendo la posizione in fase 1 (messa in mora), merge non distruttivo
(tranne --no-gestionale).

Campi dati_mora.json: data, debitore, indirizzo{via,cap,citta}, pec, debito,
spese_legali (opzionale: se vuota resta "[……..]" ed evidenziata in giallo),
output, cartella, gestionale, pdr, codice_cliente, contratto, piva, verifiche.
"""
import copy
import json
import re
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import docx
from docx.enum.text import WD_COLOR_INDEX

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE.parent / "template" / "modello_messa_in_mora.docx"
GESTIONALE = Path(r"L:\PROGETTI\COMOACQUA\GESTIONALE\gestionale_ca.json")
GIORNI_MORA = 10


# ------------------------------------------------------------ editing
def replace_in_para(p, old, new):
    """Sostituisce `old` con `new` in un paragrafo mantenendo il formattazione
    dei run. Gestisce il testo che si estende su piu' run."""
    runs = p.runs
    texts = [r.text for r in runs]
    full = "".join(texts)
    idx = full.find(old)
    if idx < 0:
        return False
    end = idx + len(old)
    pos = 0
    start_i = start_off = end_i = end_off = None
    for i, t in enumerate(texts):
        if start_i is None and idx < pos + len(t):
            start_i, start_off = i, idx - pos
        if end <= pos + len(t):
            end_i, end_off = i, end - pos
            break
        pos += len(t)
    if start_i is None or end_i is None:
        return False
    if start_i == end_i:
        runs[start_i].text = texts[start_i][:start_off] + new + texts[end_i][end_off:]
    else:
        runs[start_i].text = texts[start_i][:start_off] + new
        for i in range(start_i + 1, end_i):
            runs[i].text = ""
        runs[end_i].text = texts[end_i][end_off:]
    return True


def evidenzia(p, frammento, colore=WD_COLOR_INDEX.YELLOW):
    """Evidenzia il run che contiene `frammento` (punti da verificare)."""
    for r in p.runs:
        if frammento in r.text:
            r.font.highlight_color = colore
            return True
    return False


PEC_MAX = 40  # oltre questa lunghezza la PEC non sta nella colonna di sinistra


def impagina_destinatario(doc, debitore, pec):
    """Blocco destinatario: la colonna di destra parte dal primo tab stop dello
    stile Normal. Con nomi o PEC lunghi il testo andava a capo sul margine sinistro."""
    tabs = doc.styles["Normal"].paragraph_format.tab_stops
    if not len(tabs):
        return
    col = tabs[0].position
    paras = doc.paragraphs
    i_nome = next((i for i, p in enumerate(paras)
                   if p.text.strip() == debitore and p.text.startswith("\t")), None)
    if i_nome is None:
        return
    # ragione sociale, via e CAP: rientro sospeso, così le righe che vanno a capo
    # restano nella colonna di destra (la prima riga parte dal margine e il tab porta alla colonna)
    for p in paras[i_nome:i_nome + 3]:
        if "\t" in p.text:
            pf = p.paragraph_format
            pf.left_indent, pf.first_line_indent = col, -col

    if not pec or len(pec) <= PEC_MAX or "@" not in pec:
        return
    p_pec = paras[i_nome + 1] if i_nome + 1 < len(paras) else None
    p_cap = paras[i_nome + 2] if i_nome + 2 < len(paras) else None
    if p_pec is None or p_cap is None or not p_pec.text.startswith(pec):
        return
    # PEC lunga: la parte dopo la «@» passa all'inizio della riga di CAP e città
    run = next((r for r in p_pec.runs if pec in r.text), None)
    if run is None:
        return
    prima, dopo = pec.split("@", 1)
    run.text = run.text.replace(pec, prima + "@")
    nuovo = copy.deepcopy(run._r)
    p_cap._p.insert(p_cap._p.index(p_cap.runs[0]._r) if p_cap.runs else len(p_cap._p), nuovo)
    from docx.text.run import Run
    r2 = Run(nuovo, p_cap)
    r2.text = dopo
    # il tab iniziale della riga CAP diventa il separatore tra PEC e CAP: va già bene così


def build(data, out_path):
    doc = docx.Document(str(TEMPLATE))

    debitore = data.get("debitore", "").strip()
    if not debitore:
        sys.exit("ERRORE: campo obbligatorio mancante: debitore")
    debito = data.get("debito", "").strip()
    if not debito:
        sys.exit("ERRORE: campo obbligatorio mancante: debito")
    ind = data.get("indirizzo", {}) or {}
    pec = data.get("pec", "").strip()
    spese = data.get("spese_legali", "").strip()

    # pratiche senza piano di rientro: la frase fissa del modello sul PDR si sostituisce
    causa = data.get("frase_causa", "").strip()
    FRASE_PDR = "e per il quale non è stato rispettato il piano di pagamento sottoscritto."

    da_evidenziare = []
    for p in doc.paragraphs:
        t = p.text
        if causa and FRASE_PDR in t:
            if not replace_in_para(p, FRASE_PDR, causa):
                sys.exit("ERRORE: frase sul piano di pagamento non sostituita")
            t = p.text
        if "{data}" in t:
            replace_in_para(p, "{data}", data.get("data", date.today().strftime("%d/%m/%Y")))
        if "{debitore}" in t:
            replace_in_para(p, "{debitore}", debitore)
        if "{debito}" in t:
            replace_in_para(p, "{debito}", debito)
        if "indirizzo@pec" in t:
            if pec:
                replace_in_para(p, "indirizzo@pec", pec)
            else:
                da_evidenziare.append((p, "indirizzo@pec"))
        if "Via ……" in t:
            replace_in_para(p, "Via ……", ind.get("via", "") or "Via ……")
            if not ind.get("via"):
                da_evidenziare.append((p, "Via"))
        if "22100 Como (CO)" in t:
            nuovo = f"{ind.get('cap','')} {ind.get('citta','')}".strip()
            if nuovo and nuovo != "22100 Como (CO)":
                replace_in_para(p, "22100 Como (CO)", nuovo)
            else:
                da_evidenziare.append((p, "22100"))
        if "[……..]" in t or "[....]" in t:
            if spese:
                for seg in ("[……..]", "[....]"):
                    if seg in p.text:
                        replace_in_para(p, seg, spese)
                        break
            else:
                # spese legali lasciate in bianco: segnale per l'avvocato
                da_evidenziare.append((p, "["))

    for p, frag in da_evidenziare:
        evidenzia(p, frag)

    impagina_destinatario(doc, debitore, pec)

    doc.save(str(out_path))

    # --- controllo: non devono restare segnaposto -----------------------
    chk = docx.Document(str(out_path))
    testo = "\n".join(p.text for p in chk.paragraphs)
    for segnaposto in ("{data}", "{debitore}", "{debito}"):
        if segnaposto in testo:
            sys.exit(f"ERRORE: segnaposto non sostituito: {segnaposto}")
    if debitore.lower() not in testo.lower():
        sys.exit(f"ERRORE: il nome del debitore non compare nella lettera: {debitore}")
    if debito not in testo:
        sys.exit(f"ERRORE: l'importo non compare nella lettera: {debito}")
    return out_path


def word_export(docx_path, fmt_code, ext):
    out = docx_path.with_suffix(ext)
    ps = (
        "$w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0;"
        f"$d=$w.Documents.Open('{docx_path}');$d.SaveAs2('{out}',{fmt_code});$d.Close(0);$w.Quit()"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True)
    return out


# ------------------------------------------------------------ gestionale
def aggiorna_gestionale(data, lettera):
    """Aggiunge la posizione a gestionale_ca.json (merge non distruttivo)."""
    GESTIONALE.parent.mkdir(parents=True, exist_ok=True)
    if GESTIONALE.is_file():
        db = json.loads(GESTIONALE.read_text(encoding="utf-8"))
        if not isinstance(db.get("posizioni"), list):
            sys.exit(f"ERRORE: struttura inattesa in {GESTIONALE}")
    else:
        db = {"versione": 1, "aggiornato": 0, "posizioni": [], "comunicazioni": []}
        print(f"  (creato {GESTIONALE.name} - se hai posizioni salvate solo nel browser, "
              f"esportale e importale per non perderle)")

    oggi = date.today()
    cartella = str(data.get("cartella", ""))
    # idempotenza: se la pratica e' gia' registrata, aggiorna invece di duplicare
    esiste = next((p for p in db["posizioni"]
                   if (cartella and p.get("cartella") == cartella)
                   or (p.get("debitore", {}).get("nome", "").strip().lower()
                       == data["debitore"].strip().lower())), None)

    anno = oggi.year
    # progressivo = massimo esistente + 1 (non il conteggio: dopo una
    # cancellazione il conteggio riuserebbe un codice gia' assegnato)
    n = 0
    for p in db["posizioni"]:
        m = re.match(rf"CA-{anno}-(\d+)$", str(p.get("codice", "")))
        if m:
            n = max(n, int(m.group(1)))
    codice = f"CA-{anno}-{n+1:03d}"
    termine = (oggi + timedelta(days=GIORNI_MORA)).isoformat()
    data_iso = oggi.isoformat()

    if esiste:
        pos = esiste
        codice = pos.get("codice") or codice
        pos["importo"] = data.get("importo", pos.get("importo", ""))
        pos["cartella"] = cartella or pos.get("cartella", "")
        ind = data.get("indirizzo", {}) or {}
        deb = pos.setdefault("debitore", {})
        deb.update({k: v for k, v in {
            "nome": data["debitore"],
            "forma": data.get("forma", ""),
            "sede": (ind.get("citta") or "").strip(),
            "indirizzo": (ind.get("via") or "").strip(),
            "cap": (ind.get("cap") or "").strip(),
            "pec": data.get("pec", ""),
            "piva": data.get("piva", ""),
        }.items() if v})
        deb["sedeDivisa"] = True
        pos.setdefault("fasi", {}).setdefault("mora", {})
        pos["fasi"]["mora"].setdefault("data", data_iso)
        pos["fasi"]["mora"].setdefault("termine_data", termine)
        azione = "aggiornata"
    else:
        pos = {
            "id": "p" + datetime.now().strftime("%Y%m%d%H%M%S") + f"{n:03d}",
            "codice": codice,
            "cliente": "CA",
            "debitore": {
                "nome": data["debitore"],
                "forma": data.get("forma", ""),
                "sede": (data.get("indirizzo", {}) or {}).get("citta", ""),
                "indirizzo": (data.get("indirizzo", {}) or {}).get("via", ""),
                "cap": (data.get("indirizzo", {}) or {}).get("cap", ""),
                "sedeDivisa": True,
                "piva": data.get("piva", ""),
                "pec": data.get("pec", ""),
            },
            "annotazioni": [],
            "importo": data.get("importo", ""),
            "apertura": data_iso,
            "fase": 1,
            "stato": "attiva",
            "fasi": {
                "mora": {"data": data_iso, "termine_data": termine, "esito": ""},
                "ricorso": {"data": "", "rg": "", "autorita": "Tribunale di Como"},
                "decreto": {"data": "", "numero": "", "imm_esecutivo": False},
                "notifica": {"data": "", "esito": "", "opposizione": "", "data_opposizione": ""},
                "precetto": {"data": "", "esito": ""},
                "esecuzione": {"data": "", "tipo": "", "nota": ""},
            },
            "promemoria": {"data": termine, "testo": "Termine di adempimento messa in mora"},
            "note": "\n".join(filter(None, [
                f"Pratica aperta dalla mail: {data.get('oggetto_mail','')}",
                *[f"Pos. {q.get('n')}: cod. cliente {q.get('codice_cliente','')} - contratto "
                  f"{q.get('contratto','')} - € {q.get('importo','')}"
                  for q in (data.get("posizioni") or []) if len(data.get("posizioni") or []) > 1],
                *[f"DA VERIFICARE: {v}" for v in (data.get("verifiche") or [])],
            ])),
            "cartella": cartella,
            "eventi": [{
                "data": data_iso,
                "testo": f"Messa in mora predisposta ({lettera.name}); "
                         f"termine di adempimento {termine}",
            }],
            "storico": [{"fase": 1, "data": data_iso}],
        }
        db["posizioni"].append(pos)
        azione = "aggiunta"

    db["aggiornato"] = int(datetime.now().timestamp() * 1000)
    GESTIONALE.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
    return pos["codice"], azione, termine


# ------------------------------------------------------------ main
def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    jpath = Path(args[0]).resolve()
    data = json.loads(jpath.read_text(encoding="utf-8"))

    out = Path(data.get("output") or "messa in mora.docx")
    if not out.is_absolute():
        out = (Path(data["cartella"]) if data.get("cartella") else jpath.parent) / out

    lettera = build(data, out)
    print(f"OK  {lettera}")

    if "--doc" in args:
        print(f"OK  {word_export(lettera, 0, '.doc')}")
    if "--pdf" in args:
        print(f"OK  {word_export(lettera, 17, '.pdf')}")

    if data.get("gestionale", True) and "--no-gestionale" not in args:
        codice, azione, termine = aggiorna_gestionale(data, lettera)
        print(f"OK  gestionale: posizione {codice} {azione} "
              f"(fase 1 - mora, termine {termine}) -> {GESTIONALE}")

    print("\nPUNTI DA VERIFICARE (evidenziati in giallo nella lettera):")
    print("  - interessi e spese legali: campo lasciato in bianco [……..]")
    if data.get("verifiche"):
        for v in data["verifiche"]:
            print(f"  - {v}")
    if not data.get("verifiche"):
        print("  - nessun altro punto segnalato dall'estrazione")


if __name__ == "__main__":
    main()
