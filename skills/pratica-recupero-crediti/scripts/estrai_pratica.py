# -*- coding: utf-8 -*-
"""
estrai_pratica.py - Estrae la pratica di recupero crediti dalla mail della
cliente Como Acqua s.r.l. (CA): .msg di Outlook, .eml oppure testo della mail
letta con il connettore Gmail.

Uso:
    py estrai_pratica.py <percorso.msg|.eml> [--dest <cartella>] [--forza]
    py estrai_pratica.py --testo mail.txt --oggetto "<oggetto>" [--data gg/mm/aaaa]
                         [--allegati <cartella con gli allegati gia' scaricati>] [--dest <cartella>] [--forza]

La mail viene letta per intero da mail_ca.py: piu' posizioni (codici cliente)
nella stessa mail, indirizzo della sede distinto da quello di fornitura,
societa' in liquidazione, P.IVA cessata, piano di rientro o fatture insolute.
Le mail di sola conferma di lettura («Letto: ...») vengono scartate.

Crea/usa la cartella L:\\PROGETTI\\COMOACQUA\\ARTIFACT\\<DEBITORE>,
- salva gli allegati della mail (tranne le immagini/logo/social),
- legge l'estratto conto XLSX (fogli "Partite Contabili" e "Scadenzario"),
- estrae numero e data del Piano di Rientro dal nome dell'allegato,
- verifica che l'Importo Totale Scaduto della mail coincida con il saldo
  dell'estratto conto (si ferma se non coincide, a meno di --forza),
- scrive dati_mora.json usato da genera_mora.py.

Richiede: openpyxl; Outlook configurato su questo PC solo per i .msg (COM).
Imposta PYTHONIOENCODING=utf-8.
"""
import argparse
import json
import re
import sys
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import openpyxl

from mail_ca import fmt, forma_societa, parse_mail  # lettura completa della mail

ARTIFACT = Path(r"L:\PROGETTI\COMOACQUA\ARTIFACT")
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".ico"}


# ------------------------------------------------------------ utilita'
def eur(s):
    """'€ 24.769,85' -> Decimal('24769.85')"""
    return Decimal(str(s).replace("€", "").replace(" ", "")
                   .replace(".", "").replace(",", ".").strip())


def nome_cartella(debitore):
    """Nome cartella valido su Windows."""
    bad = '<>:"/\\|?*'
    out = "".join("_" if c in bad else c for c in debitore)
    return out.strip(" .")


# ------------------------------------------------------------ allegati
def salva_allegati(msg, dest):
    """Salva gli allegati (esclusi i logo/immagini) nella cartella pratica."""
    salvati, saltati = [], []
    usati = {}
    for i in range(1, msg.Attachments.Count + 1):
        att = msg.Attachments.Item(i)
        nome = att.FileName
        if Path(nome).suffix.lower() in IMG_EXT:
            saltati.append(nome)
            continue
        # nomi duplicati (es. piu' "image.png"): non devono succedersi
        base = Path(nome).stem
        ext = Path(nome).suffix
        if nome in usati:
            usati[nome] += 1
            nome = f"{base}_{usati[nome]}{ext}"
        else:
            usati[att.FileName] = 1
        p = dest / nome
        try:
            att.SaveAsFile(str(p))
            salvati.append(p.name)
        except Exception as e:  # allegato corrotto/non salvabile
            print(f"  AVVISO: allegato non salvato {att.FileName}: {e}")
    return salvati, saltati


def salva_allegati_eml(msg, dest):
    """Allegati di un .eml (mail scaricata da Gmail con «Scarica messaggio»)."""
    salvati, saltati = [], []
    for part in msg.iter_attachments():
        nome = part.get_filename() or "allegato"
        if Path(nome).suffix.lower() in IMG_EXT:
            saltati.append(nome)
            continue
        p = dest / nome
        i = 1
        while p.exists() and p.read_bytes() != (part.get_payload(decode=True) or b""):
            i += 1
            p = dest / f"{Path(nome).stem}_{i}{Path(nome).suffix}"
        p.write_bytes(part.get_payload(decode=True) or b"")
        salvati.append(p.name)
    return salvati, saltati


def leggi_eml(path):
    """(oggetto, corpo testo, data gg/mm/aaaa, messaggio) da un file .eml."""
    import email
    from email import policy
    from email.utils import parsedate_to_datetime
    msg = email.message_from_bytes(Path(path).read_bytes(), policy=policy.default)
    body = msg.get_body(preferencelist=("plain", "html"))
    testo = body.get_content() if body else ""
    if body is not None and body.get_content_type() == "text/html":
        testo = re.sub(r"<br\s*/?>|</p>|</div>|</li>", "\n", testo, flags=re.I)
        testo = re.sub(r"<li[^>]*>", "\n- ", testo, flags=re.I)
        testo = re.sub(r"<[^>]+>", "", testo)
        import html as _h
        testo = _h.unescape(testo)
    try:
        data = parsedate_to_datetime(msg["date"]).strftime("%d/%m/%Y")
    except Exception:
        data = ""
    return str(msg["subject"] or ""), testo, data, msg


# ------------------------------------------------------------ estratto conto
def leggi_estratto_conto(xlsx, data_rif=None):
    """Ritorna (saldo_scadenzario, saldo_partite_contabili) come Decimal.

    Nel foglio "Scadenzario" somma le righe con saldo > 0 e data di scadenza
    <= data_rif: la mail riporta l'IMPORTO TOTALE SCADUTO, quindi le bollette
    ancora in corso di scadenza (es. emesse a fine mese successivo) non
    devono essere conteggiate.
    """
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    if data_rif is None:
        data_rif = date.today()
    scad = Decimal("0")
    escluse = Decimal("0")
    if "Scadenzario" in wb.sheetnames:
        ws = wb["Scadenzario"]
        # header a riga 7: "Saldo Scadenza" e' la colonna 13
        for row in ws.iter_rows(min_row=8, values_only=True):
            if not row or row[0] is None:
                continue
            v = row[12]
            if v in (None, ""):
                continue
            try:
                val = Decimal(str(v).replace(".", "").replace(",", "."))
            except Exception:
                continue
            if val <= 0:
                continue
            scadenza = row[0]
            if hasattr(scadenza, "date") and scadenza.date() > data_rif:
                escluse += val      # bolletta non ancora scaduta
                continue
            scad += val
        if escluse:
            print(f"  scadenzario: escluse € {fmt(escluse)} di scadenze future "
                  f"(al {data_rif.strftime('%d/%m/%Y')})")
    tot = None
    if "Partite Contabili" in wb.sheetnames:
        ws = wb["Partite Contabili"]
        for row in ws.iter_rows(values_only=True):
            if row and str(row[0]).strip().upper() == "TOTALI":
                try:
                    tot = Decimal(str(row[10]).replace(".", "").replace(",", "."))
                except Exception:
                    tot = None
                break
    return scad, tot


def leggi_pdr(nomi_allegati):
    """Numero e data del piano di rientro dal nome degli allegati."""
    numero, data = "", ""
    for n in nomi_allegati:
        if not numero:
            m = re.search(r"PDR\s?\d{2}-\d{6}", n, re.IGNORECASE)
            if m:
                numero = re.sub(r"\s+", "", m.group(0)).upper()
        if not data:
            m = re.search(r"del\s?(\d{2})[-/](\d{2})[-/](\d{4})", n, re.IGNORECASE)
            if m:
                data = f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
    return numero, data


# ------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("mail", nargs="?", help="percorso del file .msg o .eml")
    ap.add_argument("--testo", help="file di testo con il corpo della mail (es. letta con il connettore Gmail)")
    ap.add_argument("--oggetto", default="", help="oggetto della mail (con --testo)")
    ap.add_argument("--data", default="", help="data della mail gg/mm/aaaa (con --testo, se manca nel testo)")
    ap.add_argument("--allegati", help="cartella (o singolo file) con gli allegati gia' scaricati (con --testo)")
    ap.add_argument("--dest", help="cartella di destinazione (default ARTIFACT/<debitore>)")
    ap.add_argument("--forza", action="store_true",
                    help="prosegue anche se l'importo della mail non coincide con l'estratto conto")
    args = ap.parse_args()

    # --- 1. corpo della mail ----------------------------------------------
    msg = eml = None
    if args.testo:
        body = Path(args.testo).read_text(encoding="utf-8", errors="replace")
        subject, data_mail = args.oggetto, args.data
    elif args.mail and args.mail.lower().endswith(".eml"):
        subject, body, data_mail, eml = leggi_eml(args.mail)
    elif args.mail and args.mail.lower().endswith(".msg"):
        msg_path = Path(args.mail).resolve()
        if not msg_path.is_file():
            sys.exit(f"ERRORE: file .msg non trovato: {msg_path}")
        try:
            import win32com.client
            ol = win32com.client.Dispatch("Outlook.Application")
            msg = ol.Session.OpenSharedItem(str(msg_path))
        except Exception as e:
            sys.exit(f"ERRORE: impossibile aprire il .msg con Outlook: {e}")
        body, subject = msg.Body, msg.Subject or ""
        data_mail = msg.ReceivedTime.strftime("%d/%m/%Y") if msg.ReceivedTime else ""
    else:
        ap.error("indicare un file .msg/.eml oppure --testo")

    # --- 2. lettura completa ----------------------------------------------
    dati = parse_mail(subject, body, data_mail)
    if dati.get("conferma_lettura"):
        if msg is not None:
            msg.Close(0)
        sys.exit("SCARTATA: e' solo una conferma di lettura («Letto: ...»), non una nuova pratica.")
    if not dati.get("debitore"):
        sys.exit("ERRORE: ragione sociale del debitore non trovata nella mail.")
    verifiche = list(dati.get("verifiche", []))

    # --- 3. cartella pratica e allegati ------------------------------------
    dest = Path(args.dest).resolve() if args.dest else ARTIFACT / nome_cartella(dati["debitore"])
    dest.mkdir(parents=True, exist_ok=True)
    print(f"Cartella pratica: {dest}")
    if msg is not None:
        try:
            salvati, saltati = salva_allegati(msg, dest)
        finally:
            try:
                msg.Close(0)
            except Exception:
                pass
    elif eml is not None:
        salvati, saltati = salva_allegati_eml(eml, dest)
    else:
        src = Path(args.allegati).resolve() if args.allegati else dest
        salvati, saltati = [], []
        elenco = sorted(src.iterdir()) if src.is_dir() else ([src] if src.is_file() else [])
        for p in elenco:
            if not p.is_file() or p.suffix.lower() in IMG_EXT or p.name in ("dati_mora.json",):
                continue
            if p.parent != dest:
                (dest / p.name).write_bytes(p.read_bytes())
            salvati.append(p.name)
        if args.testo and Path(args.testo).resolve().parent != dest:
            (dest / "mail.txt").write_text(body, encoding="utf-8")
    print(f"Allegati: {len(salvati)}" + (f" (immagini saltate: {len(saltati)})" if saltati else ""))
    for n in salvati:
        print(f"  - {n}")

    # --- 4. estratti conto (uno per posizione) -------------------------------
    xlsx = [dest / n for n in salvati if n.lower().endswith((".xlsx", ".xls"))
            and (dest / n).read_bytes()[:2] == b"PK"]
    try:
        data_rif = datetime.strptime(dati.get("data_mail", ""), "%d/%m/%Y").date()
    except Exception:
        data_rif = date.today()
    saldo_scad = Decimal("0")
    letti = 0
    for x in xlsx:
        try:
            sc, tot = leggi_estratto_conto(x, data_rif)
            print(f"Estratto conto {x.name}: saldo scadenzario € {fmt(sc)}"
                  + (f" - totali partite contabili € {fmt(tot)}" if tot is not None else ""))
            saldo_scad += sc
            letti += 1
        except Exception as e:
            verifiche.append(f"lettura estratto conto {x.name} non riuscita: {e}")
    n_pos = len(dati.get("posizioni") or []) or 1
    if not xlsx:
        verifiche.append("nessun estratto conto Excel: importo non verificato sullo scadenzario")
    elif letti < n_pos:
        verifiche.append(f"estratti conto Excel letti: {letti} su {n_pos} posizioni: verifica importo parziale")

    # --- 5. verifica importo -------------------------------------------------
    importo = eur(dati["debito"]) if dati.get("debito") else None
    if importo is not None and letti >= n_pos and saldo_scad > 0:
        if saldo_scad != importo:
            msg_err = (f"IMPORTO NON COINCIDE: mail € {fmt(importo)} != "
                       f"estratti conto € {fmt(saldo_scad)}")
            if not args.forza:
                sys.exit(f"ERRORE: {msg_err}")
            verifiche.append(msg_err + " (proseguito con --forza)")
        else:
            print(f"Verifica importo OK: € {fmt(importo)} (somma di {letti} estratti conto)")
    elif importo is not None and letti and saldo_scad == 0:
        verifiche.append("estratto conto senza foglio «Scadenzario» leggibile: "
                         "importo NON verificato, controllare a mano il TOTALE SCADUTO")

    # --- piano di rientro ---------------------------------------------
    pdr_num, pdr_data = leggi_pdr(salvati)
    pdr_num = pdr_num or dati["pdr"]["numero"]
    pdr_data = pdr_data or dati["pdr"]["data"]
    if pdr_num:
        print(f"Piano di rientro: {pdr_num} del {pdr_data or 'data non trovata'}")
    elif dati.get("causa") in ("pdr", "misto"):
        verifiche.append("la mail parla di piano di rientro ma l'allegato del PDR non e' stato individuato")

    # --- scrivi dati_mora.json ----------------------------------------
    oggi = date.today()
    dati.update({
        "output": f"messa in mora_{dati['debitore']}.docx",
        "data": oggi.strftime("%d/%m/%Y"),
        "forma": dati.get("forma") or forma_societa(dati["debitore"]),
        "pdr": {"numero": pdr_num, "data": pdr_data},
        "cartella": str(dest),
        "gestionale": True,
        "verifiche": verifiche,
    })
    out = dest / "dati_mora.json"
    out.write_text(json.dumps(dati, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK  {out}")

    # --- riepilogo ------------------------------------------------------
    print("\nRIEPILOGO")
    print(f"  Debitore : {dati['debitore']} ({dati['forma'] if 'forma' in dati else forma_societa(dati['debitore'])})")
    print(f"  Sede     : {dati['indirizzo']['via']}, {dati['indirizzo']['cap']} {dati['indirizzo']['citta']}")
    print(f"  PEC      : {dati['pec'] or '—'}")
    print(f"  P.IVA    : {dati['piva'] or '—'} ({dati['stato_piva'] or 'stato ignoto'})")
    print(f"  Debito   : € {dati['debito'] or '—'}")
    print(f"  PDR      : {pdr_num or '—'} {pdr_data or ''}")
    if len(dati.get("posizioni") or []) > 1:
        for p in dati["posizioni"]:
            print(f"  Pos. {p['n']}  : cod. {p['codice_cliente']} contr. {p['contratto']} "
                  f"{p['utenza']} {p['stato_utenza']} € {p['importo'] or '—'}")
    print(f"  Causa    : {dati.get('causa')}" + (" -> frase «fatture insolute» nella lettera" if dati.get("frase_causa") else ""))
    if verifiche:
        print("  DA VERIFICARE:")
        for v in verifiche:
            print(f"    == {v} ==")
    else:
        print("  Nessun punto da verificare.")


if __name__ == "__main__":
    main()
