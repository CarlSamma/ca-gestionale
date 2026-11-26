---
name: pratica-recupero-crediti
description: Apre una pratica di recupero crediti di Como Acqua s.r.l. dalla mail della cliente (.msg, .eml o testo letto da Gmail) (allegati: estratto conto, piano di rientro, rate, storico PEC) e genera la lettera di messa in mora in formato Word (.docx + .doc) identica al "modellomessainmora.doc" dello Studio, registrando la posizione nel gestionale. Usare quando l'utente scrive "GENERA MESSA IN MORA <percorso .msg>", "apri pratica recupero crediti <mail>" o chiede una messa in mora per un debitore di Como Acqua.
---

# Pratica di recupero crediti + messa in mora (Como Acqua)

Input: la mail di nuovo incarico dell'Ufficio Incassi di Como Acqua (inoltrata dallo Studio), come file `.msg`/`.eml` oppure letta da Gmail (vedi §0).
Output, tutto nella cartella `L:\PROGETTI\COMOACQUA\ARTIFACT\<DEBITORE>\`:
- allegati estratti dalla mail (esclusi logo/immagini/social),
- `dati_mora.json` (dati usati, modificabile e rigenerabile),
- `messa in mora_<Debitore>.docx` + `.doc` (formato dei modelli dello Studio),
- posizione aggiunta a `L:\PROGETTI\COMOACQUA\GESTIONALE\gestionale_ca.json` (fase 1 - mora).

## 0. Mail da Gmail: leggere TUTTA la mail, scartare le conferme di lettura

1. Cerca le nuove pratiche **escludendo le conferme di lettura**:
   `("Nuova pratica" OR "NUOVA PRATICA") -subject:Letto -subject:"Read:"`
   Prima di aprire ogni risultato controlla l'oggetto con `py scripts\mail_ca.py --conferma "<oggetto>"`:
   le mail «I: Letto: …», «Read: …», «Notifica di lettura» **non sono pratiche** e non vanno mostrate né elaborate.
2. Apri il messaggio completo (`get_message`, non lo snippet) e salva il **corpo testuale integrale** in
   `ARTIFACT\<DEBITORE>\mail.txt` (senza tagliare: le mail multi-posizione hanno un blocco «Dettaglio Posizione N» per ogni utenza).
3. Scarica gli allegati (estratti conto `.xlsx`, PDR, rate) nella stessa cartella.
4. Esegui:
   ```
   py scripts\estrai_pratica.py --testo "<cartella>\mail.txt" --oggetto "<oggetto>" --data gg/mm/aaaa
   ```
   Il parser (`mail_ca.py`) normalizza il testo Gmail (asterischi, righe spezzate a 72 caratteri, immagini/link),
   legge **ogni** posizione, l'indirizzo della **sede** (non quello di fornitura), PEC, P.IVA e stato, forma societaria.
5. **Verifica SEMPRE il valore di IMPORTO TOTALE SCADUTO**: lo script lo confronta con la somma delle posizioni
   indicate nella mail e con la somma degli scadenzari Excel; se non coincide si ferma (o lo segnala in `verifiche`).
   Riporta all'avvocato ogni differenza, anche di centesimi.

Segnalazioni automatiche in `verifiche`: società in liquidazione / P.IVA cessata, S.n.c. e S.a.s. (soci illimitatamente
responsabili), utenza cessata, fornitura diversa dalla sede, PEC mancante, indirizzo incompleto, «seguiranno bollette/SDI».

Test del parser (dati anonimizzati): `py -m unittest discover -s test`.

## 1. Estrai la pratica dalla mail

```
py "L:\PROGETTI\COMOACQUA\.claude\skills\pratica-recupero-crediti\scripts\estrai_pratica.py" "<percorso .msg>"
```
(per `.msg` usa Outlook via COM: richiede Outlook configurato su questo PC; un `.eml` si legge senza Outlook; testo da Gmail: vedi §0. `PYTHONIOENCODING=utf-8`).
Se la mail è solo una conferma di lettura lo script esce con «SCARTATA».

Lo script:
- crea la cartella in `ARTIFACT\<DEBITORE>` e vi salva gli allegati (salta `.png`/`.jpg` logo e social);
- legge dal corpo della mail: ragione sociale, codice cliente/contratto, tipologia utenza, indirizzo (split via/CAP/città, trattino `-` o `–`), PEC, P.IVA + stato, stato utenza, importo totale scaduto, data della mail;
- legge **tutti** gli estratti conto XLSX (uno per posizione) e per ciascuno somma le righe del foglio `Scadenzario` con saldo > 0 **e scadenza ≤ data della mail** (le bollette non ancora scadute, es. emesse a fine mese, sono escluse e vengono segnalate a video);
- **ferma il processo se l'importo della mail non coincide** con la somma degli scadenzari (usare `--forza` solo dopo aver capito la differenza e averla segnalata all'avvocato);
- estrae numero e data del PDR dal nome dell'allegato (`PDR24-001600 del 31-12-2024`) o dal testo della mail;
- se non c'è un PDR ma fatture insolute, imposta `frase_causa` automaticamente;
- scrive `dati_mora.json`.

Se manca l'estratto conto XLSX, la PEC o il PDR, lo script non si ferma: il punto finisce in `verifiche` e nella lista finale da controllare.

## 2. Controlla `dati_mora.json`

Rileggi il file e verifica: debitore (denominazione esatta come in mail), indirizzo, PEC, P.IVA, debito, `pdr`. Correggi pure a mano: ogni campo è riutilizzato alla rigenerazione. Campi:

| Campo | Note |
|---|---|
| `data` | data della lettera, `gg/mm/aaaa` (default: oggi) |
| `debitore`, `indirizzo{via,cap,citta}`, `pec`, `piva` | obbligatori |
| `debito` / `importo` | stringa formattata `24.769,85` / numero per il gestionale |
| `spese_legali` | **lasciare vuoto**: la lettera resta con `[……..]` evidenziato in giallo (deciso con l'avvocato) |
| `frase_causa` | facoltativo, per le pratiche **senza piano di rientro**: sostituisce «e per il quale non è stato rispettato il piano di pagamento sottoscritto.» (es. «come da fatture regolarmente emesse, ad oggi scadute e rimaste insolute nonostante i solleciti già inoltrati.») |
| `output` | nome file lettera |
| `gestionale` | `false` per non toccare il gestionale |
| `verifiche` | elenco punti da segnalare, stampato alla fine |

## 3. Genera la lettera

```
py "L:\PROGETTI\COMOACQUA\.claude\skills\pratica-recupero-crediti\scripts\genera_mora.py" "<cartella>\dati_mora.json" --doc
```
(`--pdf` crea anche il PDF; `--no-gestionale` non tocca il registro). Richiede `python-docx` e Microsoft Word per `.doc`/`.pdf`.

Il template è `template/modello_messa_in_mora.docx` (convertito dal `modellomessainmora.doc` dello Studio): riempie `{data}`, `{debitore}` (intestazione e oggetto), PEC + via + CAP/città, `{debito}`. Il testo fisso (letterhead, formula "Scrivo la presente…", firma avv. Baccaglini) resta invariato.

**Controllo automatico**: lo script si ferma se resta un segnaposto `{...}` o se debitore/importo non compaiono nel testo.

## 4. Aggiornamento gestionale

`genera_mora.py` crea/aggiorna `L:\PROGETTI\COMOACQUA\GESTIONALE\gestionale_ca.json` (stessa struttura dell'HTML: `versione`, `aggiornato`, `posizioni`, `comunicazioni`):
- posizione nuova: `codice` `CA-<anno>-NNN` = numero più alto esistente + 1, anagrafica con `indirizzo`/`cap`/`sede` separati (come la v1.4 del gestionale), note con le posizioni e i punti da verificare, `fase: 1`, `stato: attiva`, `fasi.mora = {data oggi, termine oggi+10, esito ""}`, `promemoria` alla scadenza, `cartella` = percorso ARTIFACT, un evento e lo storico;
- **merge non distruttivo**: se la pratica è già registrata (stessa cartella o stesso debitore) aggiorna invece di duplicare → eseguire due volte non crea doppioni.

⚠ Se il file non esisteva e l'utente ha già posizioni salvate solo nel browser (localStorage), i dati vecchi rischiano di essere sovrascritti al collegamento: fargli prima **Dati → Esporta JSON** dal gestionale, poi importarlo dopo.

## 5. Riferisci all'utente

- percorso di cartella, lettera `.docx` e `.doc`, codice posizione nel gestionale e termine di adempimento (data + 10 gg);
- **elenco punti da verificare**: spese legali in bianco `[……..]` (evidenziato in giallo), eventuali `verifiche` (allegati mancanti, disallineamenti), esito della verifica importo;
- se la verifica importo è fallita con `--forza`: riporta sempre la differenza e chiedi conferma all'avvocato prima dell'invio.

Markup nei testi di `dati_mora.json` (se servono note): `==evidenziato==` non è interpretato da questo script — usare `verifiche` per i punti da segnalare.

## Dal gestionale (v1.4)

In alternativa al §3, la lettera si può generare dal gestionale HTML: scheda posizione → fase *Messa in mora* →
**📄 Genera lettera di messa in mora**, usando il modello caricato nella scheda **📄 Modelli** (cartella `#MODELLI#`).
