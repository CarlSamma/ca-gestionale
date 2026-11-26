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
- **Scheda posizione**: barra di avanzamento a 6 fasi, anagrafica debitore con indirizzo della sede in via/CAP/comune (dalla v1.4; i dati precedenti vengono divisi automaticamente) (forma giuridica a scelta: S.r.l., S.r.l.s., S.a.s., Ditta individuale, S.p.A., S.n.c., Soc. coop., S.S.D., Persona fisica), promemoria, storico eventi (eliminazione con conferma), note, apertura della cartella pratica, pulsante *Registra fase e passa alla successiva* (alla fase 6 chiude la procedura) e, per correggere un errore, *↩ Annulla «fase» completata* che riporta alla fase precedente mantenendo i dati inseriti.
- **Stati**: attiva / sospesa / chiusa-incassata / chiusa-abbandonata.
- **Comunicazioni**: note sulle comunicazioni con la cliente CA o con la controparte; dalla scheda posizione si aprono **solo** le comunicazioni collegate a quella posizione.
- **Colonna Note** nel registro (v1.4): note manuali per posizione, ciascuna con data facoltativa; con *avvisami alla data* la nota diventa una scadenza (dashboard, colore, pannello «📌 Note con data») finché non viene segnata come fatta. La ricerca trova anche il testo delle note.
- **📄 Modelli** (v1.4): area per i modelli Word `.docx` (messa in mora, ricorso 633 c.p.c., altro). I modelli restano nel browser e, con la cartella collegata, nella sottocartella `#MODELLI#` (i `.docx` copiati lì a mano vengono importati da soli). Per ogni modello si vedono i segnaposto riconosciuti.
- **Genera lettera di messa in mora** (v1.4): usa come predefinito il modello della skill (`.claude\skills\pratica-recupero-crediti\template\modello_messa_in_mora.docx`, copiato in `#MODELLI#` e pubblicato accanto alla versione online); dalla fase *Messa in mora* della scheda; compila il modello con data, debitore, PEC, indirizzo (via, CAP, comune), importo e causa (piano di rientro non rispettato / fatture insolute). Le spese legali vuote restano `[……..]` evidenziate in giallo. Con la cartella delle posizioni collegata la lettera viene **salvata direttamente in `0-CORRISPONDENZA`** della pratica (senza sovrascrivere file esistenti). Alla fine mostra i punti da verificare (società in liquidazione, S.n.c./S.a.s., PEC o indirizzo mancanti) e registra l'evento.
- **📥 Posta CA** (solo nella versione pubblicata su claude.ai, con il connettore Gmail): legge le mail inoltrate alla casella del gestionale e permette di creare una posizione o registrare la mail su una posizione esistente; il modulo è già compilato con i dati della mail (sede, PEC, P.IVA, codici, tipologia e stato dell'utenza, importo). Le **conferme di lettura** («Letto: …») e le **ricevute PEC** di accettazione/consegna non vengono mostrate; per ogni mail di nuova pratica legge tutte le posizioni e segnala se l'**importo totale scaduto** non coincide con la somma delle posizioni.
- **📂 Cartelle delle posizioni** (v1.5, versione sul PC con Chrome/Edge): *Dati ▾ → Collega cartella POSIZIONI* collega la cartella con una sottocartella per debitore (es. `R:\Como Acqua\POSIZIONI\<DEBITORE>`). Vengono lette **solo le cartelle direttamente in POSIZIONI**: quelle tra `#` (`#CON DECRETI EMESSI#`, `#CONCORSUALE#`, `#DEFINITI-ARCHIVIATO#`, `#MODELLI#`) sono ignorate. Il pulsante **🔄 Aggiorna dalle cartelle** (registro) o **🔄 Aggiorna dalla cartella** (scheda) per ogni posizione:
  - abbina la cartella alla posizione (percorso, nome della cartella o denominazione) e propone di aprire le cartelle che non hanno ancora una posizione;
  - legge la **prima mail di CA** in `0-CORRISPONDENZA` (`.msg` di Outlook o `.eml`, escluse ricevute PEC e risposte) e compila l'**Anagrafica debitore** (denominazione, forma, via/CAP/comune, P.IVA e suo stato, PEC, codici cliente, codici contratto, tipologia e stato dell'utenza), l'**Importo Totale Scaduto**, contratti e stato delle utenze; la mail finisce nelle Comunicazioni;
  - legge l'**estratto conto** `.xlsx` (cartella `3-ESTRATTO CONTO` o allegato alla mail): le bollette scadute alla data della situazione diventano fatture, quelle non ancora scadute sono escluse, e il totale viene confrontato con l'importo della mail;
  - segnala i **documenti nuovi** dall'ultimo aggiornamento (registrati negli eventi);
  - **controlla la messa in mora** (v1.5.2): legge le lettere `.docx` dello Studio (data «Como, gg/mm/aaaa» e «termine di N giorni») e le ricevute PEC (`CONSEGNA`, `AVVISO DI MANCATA CONSEGNA`) in tutta la cartella, abbina la lettera più recente al suo invio e confronta con il gestionale: data di invio = data della ricevuta di consegna, giorni = quelli della lettera, termine = consegna + giorni. Se la PEC non è stata consegnata, o è stata consegnata a un indirizzo diverso dalla PEC del debitore, propone l'esito «PEC non consegnata — da reinviare» (nessun termine; in Scadenze compare il promemoria di reinvio). La correzione è già spuntata; dopo averla vista non viene riproposta finché i file della cartella non cambiano;
  - non sovrascrive mai dati già presenti: le differenze vengono mostrate e si sceglie quali sostituire.
- **Messa in mora** (v1.6): nella colonna *Fase* le pratiche in F1 hanno i tag **[pec] [racc] [pec+racc]** per il mezzo di invio (clic per sceglierlo, di nuovo per toglierlo; anche come «Mezzo di invio» nella scheda). L'esito può essere **«PEC non consegnata»**: nessun termine e promemoria di reinvio in Scadenze. Nella lettera generata il «termine di 10 (dieci) giorni» del modello diventa quello della pratica (es. «termine di 7 (sette) giorni»), modificabile nella finestra di generazione.
- **Colonna Codice** (v1.6.1): mostra in evidenza i **codici cliente di Como Acqua** (es. C702, C056; anche più di uno) presi dall'anagrafica o, se mancano, dal nome della mail di CA in `0-CORRISPONDENZA` («Codice Cliente_ C702 - …», «DEBITORE 013023000807 - …»); sotto, il codice interno CA-AAAA-NNN. L'ordinamento per codice segue il codice cliente.
- **📝 Genera rapporto** (v1.6): si scelgono una o più pratiche (dal registro o dal pulsante *Rapporto* nella scheda) e il periodo; il gestionale prepara un breve aggiornamento per pratica con credito, fase, stato della messa in mora, attività recenti (eventi e comunicazioni, esclusi quelli tecnici), prossimo passo e, a scelta, le note interne. Il testo si corregge nella finestra e si copia o si scarica in Word.
- **Colonna Stato** (v1.5): menu a tendina nella riga con *in attesa da CA*, *in attesa da SB*, *in attesa esito deposito*, *in attesa esito notifica*, con filtro dedicato. Lo stato della posizione (attiva, sospesa, chiusa) resta nella scheda.
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
skills/pratica-recupero-crediti/   # skill di Claude: apertura pratica dalla mail e lettera di messa in mora
  SKILL.md
  scripts/mail_ca.py               # parser delle mail di nuovo incarico (Gmail/.eml/.msg), esclude conferme di lettura e ricevute PEC
  scripts/estrai_pratica.py        # crea la cartella pratica, verifica l'importo sugli estratti conto
  scripts/genera_mora.py           # compila la lettera e aggiorna gestionale_ca.json
  test/                            # test del parser con mail anonimizzate
README.md
```

I modelli Word dello Studio e i dati reali (`gestionale_ca*.json`, mail, estratti conto) **non** sono nel repository.

## Avvertenza

I termini sono calcolati automaticamente **a solo scopo di monitoraggio**: verificarli sempre sul testo vigente del codice e sull'atto notificato.
