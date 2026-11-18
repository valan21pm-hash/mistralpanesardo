# Gestione Calcetto Pro — App installabile + Foglio Google condiviso

App per la gestione della squadra di calcio a 5: saldi, quote di iscrizione,
convocazioni, amichevoli con formazioni, referti, archivio risultati, ricariche,
spese comuni.

- **Installabile sul cellulare** come app (PWA), con icona sulla schermata home.
- **Lavora in autonomia** (nessuna AI, nessun modello LLM): solo la tua gestione.
- **Collegamento al Foglio Google in lettura e scrittura**: i dati restano nel
  foglio condiviso, quindi **due persone possono usarla insieme** usando lo stesso URL.

## Struttura

| File | Dove va | A cosa serve |
|---|---|---|
| `Codice.gs` | script.google.com (Apps Script) | Tutta la logica + Foglio Google (identico all'originale) + API per l'app |
| `index.html` | hosting web (GitHub Pages, Netlify...) | l'app vero e proprio, installabile |
| `manifest.webmanifest`, `sw.js`, `icons/` | stessa cartella dell'app | installazione PWA e funzionamento offline dell'interfaccia |

Questa app sostituisce la vecchia `Index.html` di Apps Script: **caricalo comunque**
in Apps Script solo se vuoi tenere la pagina web di riserva, non serve per l'app.

## Installazione (una volta sola, ~5 minuti)

1. **Foglio**: apri `Codice.gs`, metti l'ID del tuo foglio in `FOGLIO_ID` e incollalo
   in [script.google.com](https://script.google.com) al posto del vecchio `Codice.gs`.
2. **Pubblica** → Distribuisci → Nuova distribuzione → App web
   - Esegui come: **Me**
   - Chi ha accesso: **Chiunque abbia il link**
   - Copia l'URL che termina con **`/exec`**.
3. **App sul telefono**: pubblica la cartella su un hosting HTTPS gratuito
   (GitHub Pages va benissimo), apri l'URL con Chrome (Android) o Safari (iPhone):
   - Android: menu ⋮ → "Installa app" / "Aggiungi a schermata Home"
   - iPhone: Condividi → "Aggiungi alla schermata Home"
4. **Collegamento**: apri l'app → ⚙️ Impostazioni → incolla l'URL `/exec` →
   "Salva URL e Connetti". Vedrai subito i saldi del foglio.

**Per la seconda persona**: ripeti solo i punti 3 e 4 (installazione + stesso URL).
I dati restano unici perché stanno nel Foglio Google: chi scrive, scrive per tutti.

Ogni installazione ricorda l'URL sul telefono (niente login ogni volta).

## Uso

1. **Gestione Rosa**: aggiungi i giocatori (nascono "inattivi").
2. **Gestione Iscrizioni**: attiva i giocatori e registra le quote di iscrizione.
3. **Crea Convocazione**: partita, data, ora, campo, costo e convocati con ruolo;
   il pulsante copia l'invito formattato per WhatsApp (Allenatore e Tifoso non
   pagano la quota).
4. **Crea Amichevole**: rosa interna + giocatori esterni, poi assegni le squadre.
5. **Inserisci Referto**: risultato, presenti, gol/assist/cartellini/gol subiti e
   chi paga; la chiusura aggiorna saldi e statistiche sul foglio.
6. **Archivio Risultati**: referto WhatsApp, modifica (ricalcola i saldi) o
   eliminazione (storna tutto).
7. **Registra Ricarica Conto** e **Dividi Spesa Comune**: aggiornano i saldi.
8. **Impostazioni**: URL del foglio, nome squadra, copia dei dati in JSON.

## Note tecniche

- Le modifiche raggiungono il foglio tramite la Web App: il client invia
  `POST {fn, args}` (Content-Type `text/plain` per evitare il preflight CORS);
  se la rete blocca il POST, ricade su JSONP in GET. Nessuna chiave API.
- Il service worker mette in cache solo l'interfaccia: i dati vengono sempre
  letti/scritti sul foglio live, quindi i due utenti restano sincronizzati.
- Cambiare qualcosa nel foglio a mano? L'app legge i dati freschi ad ogni azione.
