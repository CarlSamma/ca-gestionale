# Gestionale Recupero Crediti — Cliente CA

Gestionale semplice per lo Studio Legale: organizzazione e monitoraggio delle posizioni affidate al recupero crediti per la cliente **CA (Como Acqua s.r.l.)**.

Un unico file, nessuna dipendenza, nessuna installazione:

```
gestionale_recupero_crediti.html
```

Aprire con **Chrome o Edge** (doppio clic sul file).

## Le 6 fasi della procedura

| # | Fase | Campi | Termine automatico |
|---|------|-------|--------------------|
| 1 | Messa in mora | data invio, termine di adempimento, esito (silenzio / riscontro / pagamento) | termine indicato |
| 2 | Ricorso per ingiunzione | data deposito, RG, autorità (G.d.P. fino a € 10.000, oltre Tribunale) | — |
| 3 | Decreto ingiuntivo | data, numero, "immediatamente esecutivo" | — |
| 4 | Notifica del decreto | data, esito, opposizione | **40 gg** se decreto immediatamente esecutivo, **30 gg** negli altri casi (art. 640 c.p.c.) |
| 5 | Precetto | data, esito | **10 gg** per pagare, **90 gg** di efficacia per il pignoramento (artt. 480, 481, 497 c.p.c.) |
| 6 | Azione esecutiva | data, tipo, udienza | — |

## Funzionalità

- **Registro posizioni**: dashboard (posizioni attive, importo esposto, scadenze superate, scadenze entro 7 giorni, prossima scadenza), ricerca, filtri per fase e stato, ordinamento, scadenze evidenziate in rosso/ambra.
- **Scheda posizione**: barra di avanzamento a 6 fasi, anagrafica debitore, promemoria, storico eventi, note, apertura della cartella pratica, pulsante *Registra fase e passa alla successiva*.
- **Stati**: attiva / sospesa / chiusa-incassata / chiusa-abbandonata.
- **Comunicazioni**: note personali sulle comunicazioni intercorse con la cliente CA o con la controparte (interlocutore, soggetto, canale, posizione collegata), con filtri e ricerca.
- **Stampa** del registro.

## Salvataggio dei dati

- salvataggio **automatico** nel browser (localStorage) a ogni modifica;
- **Dati ▾ → Collega cartella**: i dati vengono tenuti anche in un file `gestionale_ca.json` nella cartella scelta (Chrome/Edge, File System Access API);
- **Esporta / Importa JSON** per backup e trasferimenti.

## Struttura del repository

```
gestionale_recupero_crediti.html   # applicazione (single file: HTML + CSS + JS)
README.md
```

## Avvertenza

I termini sono calcolati automaticamente **a solo scopo di monitoraggio**: verificarli sempre sul testo vigente del codice e sull'atto notificato.
