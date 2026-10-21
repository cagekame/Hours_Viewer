# Engineering Hours Viewer — Manuale Utente
**Versione 1.0 | Flowserve Desio**

---

## 1. Panoramica dell'interfaccia

### 1.1 Barra dell'intestazione
Mostra il nome dell'applicazione, il filtro attivo (risorsa o gruppo) e il numero di record caricati.

### 1.2 Barra degli strumenti
Mostra la data dell'ultimo aggiornamento e il pulsante **⟳ Refresh from DB**. Un banner giallo appare automaticamente quando sul server sono disponibili dati più recenti.

### 1.3 Pannello filtri

| Filtro | Descrizione |
|---|---|
| **Group** | Filtra per reparto/team. Vengono mostrati solo i gruppi attivi negli ultimi 60 giorni. |
| **Resource** | Filtra per persona. Si aggiorna automaticamente quando si seleziona un gruppo. Vengono mostrate solo le risorse attive negli ultimi 60 giorni. |
| **Job** | Ricerca libera sul numero di commessa/ordine. |
| **From / To** | Intervallo di date. Clicca 📅 per aprire il calendario. |
| **Search** | Applica i filtri e aggiorna la tabella. |
| **Reset** | Ripristina i filtri ai valori predefiniti (anno corrente, tutti i gruppi). |
| **Export XLSX** | Esporta i risultati correnti in un file Excel. |

### 1.4 Tabella risultati

| Colonna | Descrizione |
|---|---|
| **Resource** | Nome della risorsa |
| **Job** | Numero commessa |
| **Phase** | Fase dell'attività |
| **Description** | Descrizione breve della commessa |
| **Notes** | Note libere inserite dalla risorsa |
| **Date** | Data in cui è stato svolto il lavoro |
| **Created On** | Data di inserimento del record nel sistema |
| **Hours** | Ore registrate |

L'ultima riga mostra sempre il **totale ore** per la selezione corrente. Clicca su un'intestazione di colonna per ordinare.

La barra di riepilogo sotto la tabella mostra:
```
N record  |  Total hours: X.Xh  |  Expected: Xh  |  Period: YYYY-MM-DD → YYYY-MM-DD
```

---

## 2. Aggiornamento dati

Clicca **⟳ Refresh from DB** per scaricare i dati più recenti dal server. Il processo richiede circa 25–30 secondi.

Un banner giallo appare automaticamente quando i dati sul server sono più recenti della cache locale:

```
⚠  Updated data available on server (Excel modified Xh ago)
```

Clicca **Update now** nel banner oppure usa il pulsante Refresh.

> **Nota:** Ogni utente ha la propria cache locale. Il refresh sul tuo PC non influisce sugli altri utenti.

---

## 3. Controllo ore

Quando si seleziona una risorsa specifica e un intervallo di date, l'applicazione verifica automaticamente se sono state registrate le ore corrette per ogni giorno lavorativo.

**Ore attese per giorno (Italia):**
- Da lunedì a giovedì: **8 ore**
- Venerdì: **6 ore**
- Weekend e festività italiane: **non conteggiati**

Il risultato viene mostrato nella barra di controllo sotto i filtri:

| Stato | Messaggio |
|---|---|
| ✅ OK | `✓  All hours complete for the selected period` |
| ⚠️ Mancanti | `⚠  N days with missing hours  (X.Xh total)` |

### 3.1 Visualizzare e escludere giorni

Quando vengono rilevate ore mancanti, clicca **View Details** per aprire la finestra di dettaglio. Per ogni giorno vengono mostrate le ore attese, le ore registrate e le ore mancanti.

**Per escludere un giorno** (es. ferie, malattia, permesso, cds):

1. Clicca **View Details** per aprire la finestra delle ore mancanti
2. Clicca sulla colonna **Exclude** (`☐`) accanto al giorno — diventa `☑` e si colora di grigio
3. L'esclusione viene salvata immediatamente e ricordata alla prossima apertura

> **Nota:** Le esclusioni sono personali e non influiscono sugli altri utenti.

---

## 4. Contatti e supporto

Per qualsiasi problema o domanda, contatta:

**Marino Caggianelli** — mcaggianelli@flowserve.com — Flowserve Desio Engineering — WORTHINGTON