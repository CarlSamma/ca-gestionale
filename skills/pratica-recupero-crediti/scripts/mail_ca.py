# -*- coding: utf-8 -*-
"""
mail_ca.py - Lettura completa delle mail di Como Acqua (nuove pratiche di
recupero crediti) e riconoscimento delle conferme di lettura.

Funziona con:
- il testo semplice restituito dal connettore Gmail (con *grassetti*,
  righe spezzate a 72 caratteri, [image: ...], intestazioni di inoltro
  «*Da:* / *Oggetto:*»);
- il corpo di un .msg di Outlook o di un .eml;
- il testo incollato (elenchi con «•»).

Uso da riga di comando (verifica rapida di una mail salvata in testo):
    py mail_ca.py mail.txt [--oggetto "..."]          -> stampa i dati estratti (JSON)
    py mail_ca.py --conferma "I: Letto: Codice Cliente ..."  -> dice se e' una conferma di lettura

Nessuna dipendenza esterna.
"""
import json
import re
import sys
from datetime import date
from decimal import Decimal, InvalidOperation

MESI = {"gennaio": 1, "febbraio": 2, "marzo": 3, "aprile": 4, "maggio": 5, "giugno": 6,
        "luglio": 7, "agosto": 8, "settembre": 9, "ottobre": 10, "novembre": 11, "dicembre": 12}

# ------------------------------------------------------------ conferme di lettura
# prefissi di inoltro / risposta che precedono l'oggetto vero (anche ripetuti: «I: R: Letto: ...»)
_PREFISSI = r"(?:(?:I|R|RE|RIF|FW|FWD|TR|INOLTRO|AW|WG|SV|VS)\s*:\s*)*"
_LETTURA = (r"(?:Letto|Non letto|Read|Not read|Lu|Gelesen|Le[ií]do|Lido|"
            r"Notifica di lettura|Conferma di lettura|Ricevuta di (?:avvenuta )?lettura|"
            r"Read receipt|Return receipt|Read-Receipt|Disposition notification)\s*:?")
RE_CONFERMA_OGGETTO = re.compile(r"^\s*" + _PREFISSI + _LETTURA, re.IGNORECASE)
RE_CONFERMA_CORPO = re.compile(
    r"(?:Oggetto|Subject)\s*:\s*\**\s*" + _PREFISSI + r"(?:Letto|Read)\s*:"
    r"|(?:il|your) messag(?:gio|e)\b.{0,120}\b(?:è stato letto|was read)"
    r"|\bè stato letto (?:il|in data)\b",
    re.IGNORECASE | re.DOTALL)


def is_conferma_lettura(oggetto, corpo=""):
    """True se la mail e' solo una conferma di lettura («Letto: ...», «Read: ...»).

    Si guarda prima l'oggetto (anche dopo «I:», «R:», «Fwd:»...). Se l'oggetto non basta,
    un corpo breve senza dati di pratica che riporta un oggetto «Letto: ...» o la frase
    «... è stato letto» conta come conferma."""
    if RE_CONFERMA_OGGETTO.match(oggetto or ""):
        return True
    c = corpo or ""
    if len(c) < 3000 and not re.search(r"Ragione Sociale|Importo (?:Totale )?Scaduto", c, re.I):
        return bool(RE_CONFERMA_CORPO.search(c))
    return False


# ------------------------------------------------------------ normalizzazione del testo
def normalizza(corpo):
    """Testo pulito, una voce per riga: toglie grassetti, immagini e link, unisce le righe
    spezzate dall'a-capo automatico e porta ogni «Etichetta: valore» su una riga sola."""
    t = (corpo or "").replace("\r\n", "\n").replace("\r", "\n").replace(" ", " ").replace("\t", " ")
    t = re.sub(r"\[image:[^\]]*\]", " ", t, flags=re.S)          # [image: logo.png] anche su piu' righe
    t = re.sub(r"<(?:https?|mailto):[^>]*>", " ", t)               # <https://...>
    t = t.replace("*", "")                                         # grassetti del testo semplice di Gmail
    t = re.sub(r"[ ]{2,}", " ", t)
    righe = []
    titolo = re.compile(r"^(?:Dettaglio Posizione\s*\d+|Riepilogo\b.*|Dati Anagrafici\b.*|Posizione\s*\d+\s*\(.*\))\s*$", re.I)
    for par in re.split(r"\n\s*\n", t):                            # paragrafi separati da righe vuote
        voce, prec = "", ""
        for r in par.split("\n"):
            r = r.strip()
            if not r:
                continue
            puntata = re.match(r"^(?:[-•·▪◦●]|\d+[.)])\s+", r)
            etichetta = re.match(r"^[A-ZÀ-Ü][^:\n]{1,60}:\s", r)
            intestazione = re.match(r"^(?:Da|Inviato|A|Cc|Ccn|Oggetto|Priorità|Importanza|From|Sent|To|Subject|Date|Data)\s*:", r, re.I)
            # riga spezzata dall'a-capo automatico (testo semplice di Gmail, ~72 caratteri):
            # la precedente era lunga o finiva con un segno che chiede un seguito
            continua = voce and not puntata and not titolo.match(r) and not titolo.match(prec) and (
                len(prec) >= 60 or prec.endswith(("€", ":", ",", "–", "-", "/", "&", "(")))
            if voce and (not continua or intestazione or etichetta and len(prec) < 60):
                righe.append(voce)
                voce = ""
            prec = r
            r = re.sub(r"^[-•·▪◦●]\s+", "", r)
            voce = (voce + " " + r).strip() if voce else r
        if voce:
            righe.append(voce)
        righe.append("")
    return "\n".join(righe)


def eur(s):
    """'€ 24.769,85' / '€.3.997,34' -> Decimal('24769.85'); None se non interpretabile."""
    m = re.search(r"(\d{1,3}(?:\.\d{3})*(?:,\d{1,2})|\d+(?:,\d{1,2})?)", str(s or ""))
    if not m:
        return None
    try:
        return Decimal(m.group(1).replace(".", "").replace(",", "."))
    except InvalidOperation:
        return None


def fmt(d):
    s = f"{d:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _campo(testo, etichetta, tutti=False):
    """Valore di «Etichetta: valore» (una voce per riga dopo normalizza)."""
    rx = re.compile(r"^\s*" + etichetta + r"\s*:\s*(.+)$", re.IGNORECASE | re.MULTILINE)
    if tutti:
        return [m.group(1).strip() for m in rx.finditer(testo)]
    m = rx.search(testo)
    return m.group(1).strip() if m else ""


def parse_indirizzo(ind):
    """'VIA A.BOITO, 14 – 20017 – RHO' -> (via, cap, citta). Trattino corto o lungo."""
    ind = (ind or "").strip().rstrip(".")
    m = re.match(r"^(?P<via>.+?)\s*[-–—,]\s*(?P<cap>\d{5})\s*[-–—]?\s*(?P<citta>.+)$", ind)
    if m:
        return m.group("via").strip(" ,"), m.group("cap"), m.group("citta").strip()
    return ind, "", ""


def forma_societa(nome):
    """Forma giuridica nella denominazione, nella grafia del gestionale."""
    u = (nome or "").upper()
    sigla = lambda *l: r"(?<![A-Z])" + r"\.?\s?".join(l) + r"\.?(?![A-Z])"
    for rx, forma in ((sigla("S", "R", "L", "S"), "S.r.l.s."), (sigla("S", "R", "L"), "S.r.l."),
                      (sigla("S", "P", "A"), "S.p.A."), (sigla("S", "A", "S"), "S.a.s."),
                      (sigla("S", "N", "C"), "S.n.c."), (sigla("S", "S", "D"), "S.S.D."),
                      (r"SOC(?:IETA'?|IETÀ)?\.?\s*COOP|COOPERATIVA", "Soc. coop.")):
        if re.search(rx, u):
            return forma
    return ""


def data_inoltro(testo):
    """Data della mail originale dall'intestazione di inoltro («Inviato: mercoledì 25 novembre 2026 16:54»)."""
    m = re.search(r"(?:Inviato|Data|Sent|Date)\s*:\s*(?:\w+,?\s+)?(\d{1,2})\s+([a-zà]+)\s+(\d{4})", testo, re.I)
    if m and m.group(2).lower() in MESI:
        return f"{int(m.group(1)):02d}/{MESI[m.group(2).lower()]:02d}/{m.group(3)}"
    m = re.search(r"(?:Inviato|Data|Sent|Date)\s*:\s*(?:\w+,?\s+)?(\d{1,2})/(\d{1,2})/(\d{4})", testo, re.I)
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"
    return ""


def oggetto_originale(testo, oggetto=""):
    """Oggetto della mail di Como Acqua: quello dell'intestazione di inoltro se c'e', altrimenti
    l'oggetto ricevuto senza prefissi «I:», «Fwd:»..."""
    m = re.search(r"^\s*(?:Oggetto|Subject)\s*:\s*(.+)$", testo, re.I | re.M)
    o = m.group(1).strip() if m else (oggetto or "")
    return re.sub(r"^\s*" + _PREFISSI, "", o).strip()


# ------------------------------------------------------------ lettura della pratica
def parse_mail(oggetto, corpo, data_mail=""):
    """Legge TUTTA la mail della nuova pratica e restituisce un dizionario con i dati,
    le posizioni (una per codice cliente / utenza) e l'elenco dei punti da verificare."""
    t = normalizza(corpo)
    ogg = oggetto_originale(t, oggetto)
    verifiche = []
    d = {"oggetto_mail": ogg, "conferma_lettura": is_conferma_lettura(oggetto, corpo)}
    if d["conferma_lettura"]:
        d["verifiche"] = ["mail di sola conferma di lettura: nessuna pratica da aprire"]
        return d

    d["data_mail"] = data_inoltro(t) or data_mail or ""

    # --- debitore ----------------------------------------------------------
    deb = _campo(t, r"Ragione Sociale(?: Debitore)?") or _campo(t, r"(?:Denominazione|Debitore|Nominativo)")
    if not deb:  # mail «libere»: «NUOVA PRATICA - DEBITORE 130934 - CANTU' FITNESS S.R.L.»
        m = re.search(r"DEBITOR[EA]\s+(?:\d+\s*-\s*)?(.+)$", ogg, re.I) or \
            re.search(r"Client[ei]\s*:\s*[^-–]+[-–]\s*(.+?)\s*[–-]\s*Nuova pratica", ogg, re.I)
        deb = m.group(1).strip() if m else ""
    d["debitore"] = re.sub(r"\s{2,}", " ", deb).strip()
    d["in_liquidazione"] = bool(re.search(r"\bliquidazion", d["debitore"] + " " + t, re.I))
    d["forma"] = forma_societa(d["debitore"])

    # --- indirizzi -----------------------------------------------------------
    sede = (_campo(t, r"Indirizzo(?: della)? Sede(?: Legale)?") or _campo(t, r"Indirizzo\s*\(Sede[^)]*\)")
            or _campo(t, r"Sede(?: Legale)?") or _campo(t, r"Indirizzo"))
    forn = _campo(t, r"Indirizzo(?: di)? Fornitura")
    d["indirizzo_raw"] = sede
    via, cap, citta = parse_indirizzo(sede)
    d["indirizzo"] = {"via": via, "cap": cap, "citta": citta}
    if forn:
        d["indirizzo_fornitura"] = forn
    d["pec"] = (_campo(t, r"(?:Indirizzo )?PEC") or "").split()[0] if _campo(t, r"(?:Indirizzo )?PEC") else ""

    # --- partita IVA / codice fiscale ---------------------------------------
    piva_raw = _campo(t, r"Partita IVA(?:\s*/\s*Codice Fiscale)?") or _campo(t, r"P\.?\s?IVA")
    d["piva_raw"] = piva_raw
    m = re.match(r"\s*([A-Z0-9]{11,16})\s*(?:[-–—]\s*(.*))?$", piva_raw or "", re.I)
    d["piva"] = m.group(1) if m else ""
    d["stato_piva"] = (m.group(2) or "").strip() if m else ""
    cf = _campo(t, r"Codice Fiscale")
    if cf and cf.split()[0] != d["piva"]:
        d["codice_fiscale"] = cf.split()[0]
    if not d["piva"]:  # P.IVA citata nel testo libero (ma non quella di Como Acqua nel piè di pagina)
        cand = [x for x in re.findall(r"\b(\d{11})\b", t) if x != "03522110133"]
        if cand:
            d["piva"] = cand[0]
            verifiche.append(f"P.IVA {cand[0]} ricavata dal testo libero: verificarla in visura")

    # --- posizioni (una o piu' utenze) -----------------------------------------
    pos = []
    blocchi = re.split(r"^\s*Dettaglio Posizione\s*(\d+)\s*$", t, flags=re.I | re.M)
    if len(blocchi) > 1:
        for i in range(1, len(blocchi), 2):
            b = blocchi[i + 1]
            b = re.split(r"^\s*(?:Riepilogo|Dettaglio Posizione)", b, flags=re.I | re.M)[0]
            imp = _campo(b, r"Importo Scaduto(?: Posizione\s*\d+)?")
            pos.append({"n": int(blocchi[i]),
                        "codice_cliente": _campo(b, r"Codice Cliente"),
                        "contratto": _campo(b, r"Codice Contratto"),
                        "utenza": _campo(b, r"Tipologia(?: di)? Utenza"),
                        "stato_utenza": _campo(b, r"Stato dell[’']?\s?Utenza"),
                        "importo": fmt(eur(imp)) if eur(imp) is not None else ""})
    else:
        cc = _campo(t, r"Codice Cliente")
        if not cc:
            m = re.search(r"Codic[ei] Client[ei]\s*:\s*([A-Z0-9][A-Z0-9\- /]*?)\s*-\s", ogg, re.I) or \
                re.search(r"DEBITOR[EA]\s+(\d+)\s*-", ogg, re.I)
            cc = m.group(1).strip() if m else ""
        pos.append({"n": 1, "codice_cliente": cc, "contratto": _campo(t, r"Codice Contratto"),
                    "utenza": _campo(t, r"Tipologia(?: di)? Utenza"),
                    "stato_utenza": _campo(t, r"Stato dell[’']?\s?Utenza"), "importo": ""})
    d["posizioni"] = pos
    d["codice_cliente"] = " / ".join(p["codice_cliente"] for p in pos if p["codice_cliente"])
    d["contratto"] = " / ".join(p["contratto"] for p in pos if p["contratto"])
    d["utenza"] = " / ".join(sorted({p["utenza"] for p in pos if p["utenza"]}))
    d["stato_utenza"] = " / ".join(p["stato_utenza"] for p in pos if p["stato_utenza"])

    # codici nell'oggetto che non trovano una posizione nel corpo
    m = re.search(r"Codic[ei] Client[ei]\s*:\s*(.+?)\s+[–-]\s*Nuova pratica", ogg, re.I)
    if m:
        cod_ogg = set(re.findall(r"\b(?=[A-Z0-9\-]*\d)[A-Z0-9][A-Z0-9\-]{2,}\b", m.group(1).upper()))
        cod_pos = {p["codice_cliente"].upper() for p in pos}
        manc = sorted(c for c in cod_ogg - cod_pos if not re.fullmatch(r"\d{1,2}", c))
        if manc and len(pos) > 1:
            verifiche.append("codici cliente dell'oggetto senza dettaglio nel corpo: " + ", ".join(manc))

    # --- importo totale ------------------------------------------------------
    imp_raw = _campo(t, r"Importo Totale Scaduto(?:\s*\([^)]*\))?")
    if not imp_raw:
        m = re.search(r"Importo (?:totale )?scaduto[^€\n]{0,60}€\s*\.?\s*([\d.]+,\d{2})", t, re.I)
        imp_raw = "€ " + m.group(1) if m else ""
    m = re.match(r"^(.*?\d,\d{2}(?:\s*\([^)]*\))?)", imp_raw or "")
    d["importo_raw"] = m.group(1) if m else imp_raw
    tot = eur(imp_raw)
    somma = sum((eur(p["importo"]) or Decimal(0)) for p in pos) if any(p["importo"] for p in pos) else None
    if tot is None and somma is not None:
        tot = somma
        verifiche.append("importo totale non indicato nella mail: usata la somma delle posizioni")
    if tot is not None and somma is not None and somma != tot:
        verifiche.append(f"IMPORTO TOTALE DA VERIFICARE: la mail indica € {fmt(tot)} ma la somma delle "
                         f"posizioni e' € {fmt(somma)}")
    if tot is None:
        verifiche.append("importo totale scaduto non trovato nella mail")
    d["debito"] = fmt(tot) if tot is not None else ""
    d["importo"] = float(tot) if tot is not None else ""

    # --- piano di rientro e causa del credito -------------------------------
    m = re.search(r"(PDR\s?\d{2}-\d{6})(?:\s+del\s+(\d{1,2}[/-]\d{1,2}[/-]\d{4}))?", t, re.I)
    d["pdr"] = {"numero": re.sub(r"\s", "", m.group(1)).upper() if m else "",
                "data": (m.group(2) or "").replace("-", "/") if m else ""}
    con_pdr = bool(m) or bool(re.search(r"PDR non rispettato|piano di rientro", ogg + " " + t, re.I))
    con_fatt = bool(re.search(r"fatture (?:regolarmente emesse|insolute)|bollette (?:insolute|non)", ogg + " " + t, re.I))
    d["causa"] = "pdr" if con_pdr and not con_fatt else ("misto" if con_pdr else "fatture")
    if d["causa"] == "fatture":
        d["frase_causa"] = ("come da fatture regolarmente emesse, ad oggi scadute e rimaste insolute "
                            "nonostante i solleciti già inoltrati.")
    elif d["causa"] == "misto":
        verifiche.append("credito misto (piano di rientro non rispettato + fatture insolute): "
                         "valutare la frase della lettera sulla causa del credito")

    # --- punti da segnalare all'avvocato -------------------------------------
    if d["in_liquidazione"]:
        verifiche.append("societa' IN LIQUIDAZIONE: valutare «alla c.a. del liquidatore»")
    if re.search(r"cessat", d["stato_piva"], re.I):
        verifiche.append(f"P.IVA cessata ({d['stato_piva']}): verificare in visura la sede e il soggetto da intimare")
    if d["forma"] == "S.n.c.":
        verifiche.append("S.n.c.: soci illimitatamente responsabili (art. 2291 c.c.), valutare l'invio anche ai soci")
    if d["forma"] == "S.a.s.":
        verifiche.append("S.a.s.: valutare l'invio anche al socio accomandatario (art. 2313 c.c.)")
    if any(re.search(r"cessat", p["stato_utenza"], re.I) for p in pos):
        verifiche.append("utenza/contratto CESSATO: " + "; ".join(
            f"pos. {p['n']} {p['stato_utenza']}" for p in pos if re.search(r"cessat", p["stato_utenza"], re.I)))
    if forn and parse_indirizzo(forn)[0].upper() != via.upper():
        verifiche.append(f"indirizzo di fornitura diverso dalla sede ({forn}): la lettera va alla sede")
    if not d["pec"]:
        verifiche.append("PEC del debitore assente nella mail")
    if not via or not cap:
        verifiche.append(f"indirizzo della sede incompleto o non interpretato: «{sede}»")
    if re.search(r"Seguiranno .{0,40}bollette", t, re.I):
        verifiche.append("bollette e ricevute SDI non ancora inviate da Como Acqua: richiederle prima del ricorso")
    if not d["debitore"]:
        verifiche.append("ragione sociale del debitore non trovata")
    d["verifiche"] = verifiche
    return d


def main(argv):
    if len(argv) >= 2 and argv[0] == "--conferma":
        print("CONFERMA DI LETTURA" if is_conferma_lettura(argv[1]) else "mail normale")
        return
    if not argv:
        print(__doc__)
        return
    testo = open(argv[0], encoding="utf-8", errors="replace").read()
    ogg = argv[argv.index("--oggetto") + 1] if "--oggetto" in argv else ""
    print(json.dumps(parse_mail(ogg, testo), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main(sys.argv[1:])
