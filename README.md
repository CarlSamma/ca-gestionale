# Gestionale Recupero Crediti — Cliente CA

Gestionale semplice per lo Studio Legale: organizzazione e monitoraggio delle posizioni affidate al recupero crediti per la cliente **CA (Como Acqua s.r.l.)**.

Un unico file, nessuna dipendenza, nessuna installazione:

```
gestionale_recupero_crediti.html
```

Aprire con **Chrome o Edge** (doppio clic sul file).

## Le 6 fasi della procedura

| # | Fase | Campi | Termini calcolati automaticamente |
|---|------|-------|-----------------------------------|
| 1 | Messa in mora | data invio, giorni per adempiere (default 10), termine, esito | termine = invio + giorni (modificabile a mano). Se scade senza adempimento compare l'avviso **«Pronta per il ricorso»** |
| 2 | Ricorso per ingiunzione | data deposito, RG, autorità (G.d.P. fino a € 10.000, oltre Tribunale) | — (mostra la prova scritta: fatture e/o riconoscimento di debito) |
| 3 | Decreto ingiuntivo | data pronuncia, numero, termine di opposizione indicato nel decreto, provvisoria esecutività, notifica all'estero | **notifica entro 60 gg** dalla pronuncia (**90** se all'estero), a pena di inefficacia (art. 644 c.p.c.) |
| 4 | Notifica del decreto | data, esito, opposizione | **opposizione entro 40 gg** dalla notifica (art. 641 c.p.c.) o nel diverso termine indicato nel decreto; con decreto provvisoriamente esecutivo il precetto può essere notificato subito |
| 5 | Precetto | data, esito | fine del termine dilatorio di **10 gg** (artt. 480, 482); **90 gg** di efficacia per il pignoramento (art. 481) |
| 6 | Azione esecutiva | data, tipo, udienza | — |

### Regole di computo
- si esclude il giorno iniziale (art. 155 c.p.c.);
- **sospensione feriale 1–31 agosto** (L. 742/1969) applicata ai termini processuali degli artt. 644 e 641; non applicata ai termini del precetto;
- termine che scade di **sabato o in un giorno festivo** prorogato al primo giorno non festivo (art. 155 c. 4-5); festività nazionali incluse Pasqua e Pasquetta (il santo patrono locale non proroga i termini);
- le scadenze già assolte (es. notifica eseguita, pignoramento avviato, opposizione proposta) restano visibili in grigio ma **non contano** come ritardi.

## Funzionalità

- **Registro posizioni**: dashboard (posizioni attive, importo esposto, scadenze superate, entro 7 giorni, pronte per il ricorso, prossima scadenza), ricerca (anche per numero di fattura), filtri per fase e stato con conteggi coerenti, ordinamento, scadenze evidenziate in rosso/ambra.
- **Fatture e titolo del credito**: elenco fatture (numero, data, scadenza, importo, pagato) con totale, pagato e **residuo dovuto calcolato**; riconoscimento di debito (sì/no, data, forma). È la prova scritta mostrata nella fase del ricorso. Senza fatture si può ancora indicare un importo a mano.
- **Scheda posizione**: barra di avanzamento a 6 fasi, anagrafica debitore, promemoria, storico eventi (eliminazione con conferma), note, apertura della cartella pratica, pulsante *Registra fase e passa alla successiva* (alla fase 6 chiude la procedura).
- **Stati**: attiva / sospesa / chiusa-incassata / chiusa-abbandonata.
- **Comunicazioni**: note sulle comunicazioni con la cliente CA o con la controparte; dalla scheda posizione si aprono **solo** le comunicazioni collegate a quella posizione.
- **Stampa** del registro.
- **Codici posizione** `CA-AAAA-NNN` progressivi senza duplicati anche dopo eliminazioni.

## Salvataggio dei dati

- salvataggio **automatico** nel browser (localStorage) a ogni modifica; più schede dello stesso browser restano allineate;
- **Dati ▾ → Collega cartella**: i dati vengono tenuti anche in un file `gestionale_ca.json` nella cartella scelta (Chrome/Edge, File System Access API);
- **uso da più computer sulla stessa cartella** (rete, OneDrive…): il gestionale riconosce se il file è stato modificato da un altro computer. Se l'altro ha modificato e questo no, carica il file; se entrambi hanno modificato, **chiede quale versione tenere** e salva l'altra come copia `gestionale_ca_conflitto_….json` nella cartella. Il controllo avviene all'avvio, a ogni salvataggio e quando si torna sulla finestra;
- **Esporta / Importa JSON** per backup e trasferimenti. I file delle versioni precedenti vengono convertiti automaticamente.

## Struttura del repository

```
gestionale_recupero_crediti.html   # applicazione (single file: HTML + CSS + JS)
README.md
```

## Avvertenza

I termini sono calcolati automaticamente **a solo scopo di monitoraggio**: verificarli sempre sul testo vigente del codice e sull'atto notificato.
